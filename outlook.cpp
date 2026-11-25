#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <oleauto.h>
#include <mapix.h>
#include <mapitags.h>
#include <mapicode.h>
#include <cstdio>
#include <fcntl.h>
#include <io.h>
#include <cwchar>
#include <cstring>
#include <climits>
#include <type_traits>
#include <memory>
#include <optional>
#include <vector>
#include <cerrno>
#include <stdexcept>
#include <string>

static void check(HRESULT status, const char* operation) {
    if (FAILED(status)) {
        char message[256];
        std::snprintf(message, sizeof(message), "%s (HRESULT 0x%08lX)",
                      operation, static_cast<unsigned long>(status));
        throw std::runtime_error(message);
    }
}

static std::string utf8(const wchar_t* value, size_t length) {
    if (!length) return {};
    if (length > INT_MAX) throw std::runtime_error("Text exceeds the Windows conversion limit");
    const auto size = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS,
        value, static_cast<int>(length), nullptr, 0, nullptr, nullptr);
    if (!size) throw std::runtime_error("UTF-8 conversion failed");
    std::string text(size, '\0');
    WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value,
        static_cast<int>(length), text.data(), size, nullptr, nullptr);
    return text;
}

static void output(const wchar_t* value, size_t length) {
    const auto text = utf8(value, length);
    std::fwrite(text.data(), 1, text.size(), stdout);
}

static void line(const wchar_t* value, size_t length) {
    output(value, length);
    std::putchar('\n');
}

struct Release {
    template<class T> void operator()(T* value) const { value->Release(); }
};
template<class T> using Interface = std::unique_ptr<T, Release>;

struct Apartment {
    Apartment() { check(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED), "CoInitializeEx"); }
    ~Apartment() { CoUninitialize(); }
};

struct Value : VARIANT {
    Value() { VariantInit(this); }
    explicit Value(long number) : Value() { vt = VT_I4; lVal = number; }
    explicit Value(const wchar_t* text) : Value() {
        vt = VT_BSTR;
        bstrVal = SysAllocString(text);
        if (!bstrVal) throw std::bad_alloc();
    }
    Value(const Value&) = delete;
    Value(Value&& other) noexcept : VARIANT(other) { VariantInit(&other); }
    ~Value() { VariantClear(this); }
    IDispatch* object() const {
        if (vt != VT_DISPATCH || !pdispVal) throw std::runtime_error("Outlook returned no COM object");
        return pdispVal;
    }
    long number() const {
        if (vt != VT_I4) throw std::runtime_error("Outlook returned an unexpected numeric type");
        return lVal;
    }
    void print() const {
        if (vt != VT_BSTR) throw std::runtime_error("Outlook returned an unexpected string type");
        line(bstrVal, SysStringLen(bstrVal));
    }
};

struct DispatchException : EXCEPINFO {
    DispatchException() : EXCEPINFO{} {}
    ~DispatchException() {
        SysFreeString(bstrSource);
        SysFreeString(bstrDescription);
        SysFreeString(bstrHelpFile);
    }
};

static Value invoke(IDispatch* object, const wchar_t* member, WORD flags,
                    VARIANT* argument = nullptr, UINT argumentCount = 1) {
    DISPID id;
    auto name = const_cast<wchar_t*>(member);
    HRESULT status = object->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(status)) {
        const auto operation = "Look up Outlook member " + utf8(member, std::wcslen(member));
        check(status, operation.c_str());
    }
    DISPID propertyPut = DISPID_PROPERTYPUT;
    const bool write = (flags & DISPATCH_PROPERTYPUT) != 0;
    DISPPARAMS params{argument, write ? &propertyPut : nullptr,
                      argument ? argumentCount : 0u, write ? 1u : 0u};
    DispatchException exception;
    Value result;
    status = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                            &params, write ? nullptr : &result, &exception, nullptr);
    if (FAILED(status)) {
        auto operation = "Invoke Outlook member " + utf8(member, std::wcslen(member));
        if (status == DISP_E_EXCEPTION) {
            if (exception.pfnDeferredFillIn) exception.pfnDeferredFillIn(&exception);
            if (exception.bstrDescription)
                operation += " - " + utf8(exception.bstrDescription, SysStringLen(exception.bstrDescription));
            if (FAILED(exception.scode)) status = exception.scode;
        }
        check(status, operation.c_str());
    }
    return result;
}

static Interface<IDispatch> automation(const wchar_t* progId, DWORD context, const char* name, bool reuseRunning = false) {
    CLSID id;
    const auto resolve = std::string("Resolve ") + name;
    check(CLSIDFromProgID(progId, &id), resolve.c_str());
    IDispatch* result = nullptr;
    if (reuseRunning) {
        IUnknown* activeRaw = nullptr;
        if (SUCCEEDED(GetActiveObject(id, nullptr, &activeRaw))) {
            Interface<IUnknown> active(activeRaw);
            check(active->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&result)),
                  "Retrieve running Outlook automation object");
            return Interface<IDispatch>(result);
        }
    }
    const auto activate = std::string("Activate ") + name;
    check(CoCreateInstance(id, nullptr, context, IID_IDispatch,
                          reinterpret_cast<void**>(&result)), activate.c_str());
    return Interface<IDispatch>(result);
}

static void printCollection(IDispatch* collection, const wchar_t* member, long maximum) {
    const auto count = invoke(collection, L"Count", DISPATCH_PROPERTYGET).number();
    for (long index = 1; index <= count && index <= maximum; ++index) {
        Value key(index);
        auto item = invoke(collection, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &key);
        auto property = invoke(item.object(), member, DISPATCH_PROPERTYGET);
        property.print();
    }
}

static WORD dispatchOperation(const std::wstring& token) {
    if (token.rfind(L"get:", 0) == 0) return DISPATCH_PROPERTYGET;
    if (token.rfind(L"call:", 0) == 0) return DISPATCH_METHOD;
    if (token.rfind(L"put:", 0) == 0) return DISPATCH_PROPERTYPUT;
    return 0;
}

static Value dispatchArgument(const std::wstring& token) {
    if (token.rfind(L"s:", 0) == 0) return Value(token.c_str() + 2);
    Value value;
    if (token == L"missing") {
        value.vt = VT_ERROR;
        value.scode = DISP_E_PARAMNOTFOUND;
    } else if (token == L"true" || token == L"false") {
        value.vt = VT_BOOL;
        value.boolVal = token == L"true" ? VARIANT_TRUE : VARIANT_FALSE;
    } else if (token.rfind(L"i:", 0) == 0) {
        const wchar_t* input = token.c_str() + 2;
        wchar_t* end = nullptr;
        errno = 0;
        const long number = std::wcstol(input, &end, 10);
        if (end == input || *end || errno == ERANGE)
            throw std::runtime_error("Expected a signed 32-bit integer after i:");
        return Value(number);
    } else {
        throw std::runtime_error("Expected a COM argument prefixed with s: or i:, or true, false, or missing");
    }
    return value;
}

static void dispatchCommand(IDispatch* app, const wchar_t* const* arguments) {
    std::optional<Value> current;
    while (*arguments) {
        const std::wstring step(*arguments++);
        const WORD operation = dispatchOperation(step);
        const auto member = step.substr(step.find(L':') + 1);
        if (!operation || member.empty())
            throw std::runtime_error("Expected get:NAME, call:NAME or put:NAME");
        std::vector<Value> values;
        while (*arguments && !dispatchOperation(*arguments))
            values.push_back(dispatchArgument(*arguments++));
        if (operation == DISPATCH_PROPERTYPUT && values.size() != 1)
            throw std::runtime_error("put:NAME requires one argument");
        std::vector<VARIANT> parameters(values.rbegin(), values.rend());
        auto result = invoke(current ? current->object() : app, member.c_str(), operation,
                             parameters.empty() ? nullptr : parameters.data(),
                             static_cast<UINT>(parameters.size()));
        if (!*arguments || (result.vt != VT_EMPTY && result.vt != VT_NULL))
            current.emplace(std::move(result));
    }
    if (current->vt == VT_EMPTY || current->vt == VT_NULL) return;
    if (current->vt == VT_DISPATCH)
        throw std::runtime_error("COM returned an object. Add a step to select a value or call a method.");
    Value printable;
    check(VariantChangeType(&printable, &*current, VARIANT_ALPHABOOL, VT_BSTR),
          "Convert COM result to text");
    printable.print();
}

static void comCheck() {
    Apartment apartment;
    auto dictionary = automation(L"Scripting.Dictionary", CLSCTX_INPROC_SERVER, "Scripting.Dictionary");
    const wchar_t* steps[] = {L"put:CompareMode", L"i:1", L"call:Add",
                              L"s:check", L"s:Outlook CLI \u2713",
                              L"get:Item", L"s:CHECK", nullptr};
    dispatchCommand(dictionary.get(), steps);
}

static bool sameText(const Value& value, const wchar_t* expected) {
    return value.vt == VT_BSTR && value.bstrVal &&
           CompareStringOrdinal(value.bstrVal, -1, expected, -1, TRUE) == CSTR_EQUAL;
}

static void exchangeStatus(IDispatch* app, wchar_t** arguments) {
    auto session = invoke(app, L"Session", DISPATCH_PROPERTYGET);
    auto profile = invoke(session.object(), L"CurrentProfileName", DISPATCH_PROPERTYGET);
    if (arguments[1] && !sameText(profile, arguments[1]))
        throw std::runtime_error("Outlook is using a different profile. Open the requested profile first.");
    auto accounts = invoke(session.object(), L"Accounts", DISPATCH_PROPERTYGET);
    const auto count = invoke(accounts.object(), L"Count", DISPATCH_PROPERTYGET).number();
    for (long i = 1; i <= count; ++i) {
        Value index(i);
        auto account = invoke(accounts.object(), L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &index);
        // olExchange is zero in OlAccountType.
        if (invoke(account.object(), L"AccountType", DISPATCH_PROPERTYGET).number() != 0)
            continue;
        auto address = invoke(account.object(), L"SmtpAddress", DISPATCH_PROPERTYGET);
        if (!sameText(address, arguments[0])) continue;
        auto store = invoke(account.object(), L"DeliveryStore", DISPATCH_PROPERTYGET);
        auto primary = invoke(session.object(), L"DefaultStore", DISPATCH_PROPERTYGET);
        auto storeId = invoke(store.object(), L"StoreID", DISPATCH_PROPERTYGET);
        auto primaryId = invoke(primary.object(), L"StoreID", DISPATCH_PROPERTYGET);
        if (primaryId.vt != VT_BSTR || !primaryId.bstrVal || !sameText(storeId, primaryId.bstrVal))
            throw std::runtime_error("Connection status can only be verified for the default Exchange mailbox.");
        const long mode = invoke(session.object(), L"ExchangeConnectionMode", DISPATCH_PROPERTYGET).number();
        // OlExchangeConnectionMode distinguishes cached disconnected from connected modes.
        if (mode != 500 && mode != 600 && mode != 700 && mode != 800)
            throw std::runtime_error("The Exchange mailbox is configured, but Outlook does not report it connected.");
        std::puts("Connected Exchange mailbox");
        address.print();
        profile.print();
        return;
    }
    throw std::runtime_error("The active Outlook profile has no Exchange account matching that address.");
}

static void comCommand(IDispatch* app, const std::wstring& command, wchar_t** arguments) {
    if (command == L"com") {
        dispatchCommand(app, arguments);
        return;
    }
    if (command == L"exchange-status") {
        exchangeStatus(app, arguments);
        return;
    }
    if (command == L"sync") {
        const auto profile = arguments[0] ? std::wstring(L"s:") + arguments[0] : L"missing";
        const wchar_t* steps[] = {L"get:Session", L"call:Logon", profile.c_str(),
                                  L"missing", L"false", L"true",
                                  L"call:SendAndReceive", L"false", nullptr};
        dispatchCommand(app, steps);
        return;
    }
    if (command == L"version") {
        invoke(app, L"Version", DISPATCH_PROPERTYGET).print();
        return;
    }
    if (command == L"draft" || command == L"send") {
        Value type(0L);
        auto item = invoke(app, L"CreateItem", DISPATCH_METHOD, &type);
        const wchar_t* fields[] = {L"To", L"Subject", L"Body"};
        for (int i = 0; i < 3; ++i) {
            Value value(arguments[i]);
            invoke(item.object(), fields[i], DISPATCH_PROPERTYPUT, &value);
        }
        if (command == L"send") {
            if (!*arguments[0]) throw std::runtime_error("A recipient is required");
            auto recipients = invoke(item.object(), L"Recipients", DISPATCH_PROPERTYGET);
            auto resolved = invoke(recipients.object(), L"ResolveAll", DISPATCH_METHOD);
            if (resolved.vt != VT_BOOL || resolved.boolVal != VARIANT_TRUE)
                throw std::runtime_error("Outlook could not resolve every recipient");
        }
        invoke(item.object(), command == L"draft" ? L"Save" : L"Send", DISPATCH_METHOD);
        if (command == L"draft")
            invoke(item.object(), L"EntryID", DISPATCH_PROPERTYGET).print();
        return;
    }
    Value mapi(L"MAPI");
    auto session = invoke(app, L"GetNamespace", DISPATCH_METHOD, &mapi);
    if (command == L"folders") {
        auto folders = invoke(session.object(), L"Folders", DISPATCH_PROPERTYGET);
        printCollection(folders.object(), L"Name", LONG_MAX);
    } else {
        Value folderId(6L);
        auto inbox = invoke(session.object(), L"GetDefaultFolder", DISPATCH_METHOD, &folderId);
        auto items = invoke(inbox.object(), L"Items", DISPATCH_PROPERTYGET);
        Value received(L"[ReceivedTime]");
        VARIANT sortArguments[2]{};
        sortArguments[0].vt = VT_BOOL;
        sortArguments[0].boolVal = VARIANT_TRUE;
        sortArguments[1] = received;
        invoke(items.object(), L"Sort", DISPATCH_METHOD, sortArguments, 2);
        printCollection(items.object(), L"Subject", 50);
    }
}

struct Unload {
    void operator()(HINSTANCE module) const { FreeLibrary(module); }
};
using Module = std::unique_ptr<std::remove_pointer_t<HMODULE>, Unload>;

static Module loadMapi() {
    const DWORD needed = GetEnvironmentVariableW(L"OUTLOOK_MAPI_DLL", nullptr, 0);
    std::wstring path;
    if (needed) {
        path.resize(needed);
        const DWORD copied = GetEnvironmentVariableW(L"OUTLOOK_MAPI_DLL", path.data(), needed);
        if (!copied || copied >= needed) throw std::runtime_error("Could not read OUTLOOK_MAPI_DLL");
        path.resize(copied);
    } else {
        wchar_t programFiles[MAX_PATH];
        const DWORD length = GetEnvironmentVariableW(L"ProgramFiles", programFiles, MAX_PATH);
        if (!length || length >= MAX_PATH) throw std::runtime_error("Could not locate Program Files");
        path = std::wstring(programFiles, length) + LR"(\Microsoft Office\root\Office16\OLMAPI32.DLL)";
        if (GetFileAttributesW(path.c_str()) == INVALID_FILE_ATTRIBUTES)
            throw std::runtime_error("Outlook MAPI provider not found. Set OUTLOOK_MAPI_DLL to the Windows path of Outlook's OLMAPI32.DLL");
    }
    if (path.size() < 3 || path[1] != L':' || path[2] != L'\\')
        throw std::runtime_error("OUTLOOK_MAPI_DLL must be an absolute Windows drive path");
    Module module(LoadLibraryExW(path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (!module) check(HRESULT_FROM_WIN32(GetLastError()), "Load Outlook Extended MAPI provider");
    return module;
}

template<class Function> static Function* symbol(HMODULE module, const char* name) {
    FARPROC address = GetProcAddress(module, name);
    if (!address) throw std::runtime_error(std::string("Missing Extended MAPI export: ") + name);
    Function* result;
    static_assert(sizeof(result) == sizeof(address));
    std::memcpy(&result, &address, sizeof(result));
    return result;
}

struct Mapi {
    Module module = loadMapi();
    decltype(&MAPIUninitialize) uninitialize = symbol<decltype(MAPIUninitialize)>(module.get(), "MAPIUninitialize");
    decltype(&MAPILogonEx) logon = symbol<decltype(MAPILogonEx)>(module.get(), "MAPILogonEx");
    decltype(&MAPIFreeBuffer) freeBuffer = symbol<decltype(MAPIFreeBuffer)>(module.get(), "MAPIFreeBuffer");
    Mapi() { check(symbol<decltype(MAPIInitialize)>(module.get(), "MAPIInitialize")(nullptr), "MAPIInitialize"); }
    ~Mapi() { uninitialize(); }
    void freeRows(SRowSet* rows) const {
        for (ULONG index = 0; index < rows->cRows; ++index) freeBuffer(rows->aRow[index].lpProps);
        freeBuffer(rows);
    }
};

static void printEntryId(const SBinary& entry) {
    for (ULONG i = 0; i < entry.cb; ++i)
        std::printf("%02X", static_cast<unsigned>(entry.lpb[i]));
}

static void printTable(const Mapi& mapi, IMAPITable* table, ULONG tag, ULONG limit,
                       bool entryIds = false) {
    SizedSPropTagArray(2, columns) = {entryIds ? 2u : 1u, {tag, PR_ENTRYID}};
    check(table->SetColumns(reinterpret_cast<LPSPropTagArray>(&columns), 0), "Set MAPI table columns");
    for (ULONG total = 0; total < limit;) {
        SRowSet* raw = nullptr;
        const auto batch = static_cast<LONG>(limit - total > 50 ? 50 : limit - total);
        check(table->QueryRows(batch, 0, &raw), "Query MAPI table rows");
        auto cleanup = [&mapi](SRowSet* rows) { mapi.freeRows(rows); };
        const std::unique_ptr<SRowSet, decltype(cleanup)> rows(raw, cleanup);
        if (!rows || !rows->cRows) break;
        total += rows->cRows;
        for (ULONG index = 0; index < rows->cRows; ++index) {
            const auto& row = rows->aRow[index];
            if (entryIds) {
                if (row.cValues == 2 && row.lpProps[1].ulPropTag == PR_ENTRYID)
                    printEntryId(row.lpProps[1].Value.bin);
                std::putchar('\t');
            }
            if (row.cValues && row.lpProps[0].ulPropTag == tag && row.lpProps[0].Value.lpszW) {
                const auto* value = row.lpProps[0].Value.lpszW;
                line(value, std::wcslen(value));
            } else {
                std::putchar('\n');
            }
        }
    }
}

static SBinary binaryProperty(const Mapi& mapi, IMAPIProp* object, ULONG tag,
                               SPropValue*& properties) {
    SizedSPropTagArray(1, columns) = {1, {tag}};
    ULONG count = 0;
    check(object->GetProps(reinterpret_cast<LPSPropTagArray>(&columns), 0, &count, &properties),
          "Read MAPI property");
    if (!properties || count != 1 || properties[0].ulPropTag != tag) {
        if (properties) mapi.freeBuffer(properties);
        properties = nullptr;
        throw std::runtime_error("MAPI did not return the required entry ID");
    }
    return properties[0].Value.bin;
}

static Interface<IMAPIFolder> openFolder(IMsgStore* store, ULONG size, LPENTRYID id, ULONG flags) {
    ULONG kind = 0;
    IUnknown* raw = nullptr;
    check(store->OpenEntry(size, id, nullptr, flags, &kind, &raw), "Open MAPI folder");
    Interface<IUnknown> object(raw);
    if (kind != MAPI_FOLDER || !object)
        throw std::runtime_error("MAPI returned an unexpected folder type");
    return Interface<IMAPIFolder>(static_cast<IMAPIFolder*>(object.release()));
}

static void setMessageProperties(const Mapi& mapi, IMessage* message, ULONG count,
                                 SPropValue* properties) {
    SPropProblemArray* raw = nullptr;
    check(message->SetProps(count, properties, &raw), "Set MAPI message properties");
    auto freeProblems = [&mapi](SPropProblemArray* problems) { mapi.freeBuffer(problems); };
    const std::unique_ptr<SPropProblemArray, decltype(freeProblems)> problems(raw, freeProblems);
    if (problems && problems->cProblem)
        check(problems->aProblem[0].scode, "Set MAPI message property");
}

static Interface<IMessage> createNote(const Mapi& mapi, IMAPIFolder* folder,
                                      wchar_t* subject, wchar_t* body) {
    IMessage* raw = nullptr;
    check(folder->CreateMessage(nullptr, 0, &raw), "Create MAPI message");
    Interface<IMessage> message(raw);
    wchar_t messageClass[] = L"IPM.Note";
    SPropValue properties[4]{};
    properties[0].ulPropTag = PR_MESSAGE_CLASS_W;
    properties[0].Value.lpszW = messageClass;
    properties[1].ulPropTag = PR_SUBJECT_W;
    properties[1].Value.lpszW = subject;
    properties[2].ulPropTag = PR_BODY_W;
    properties[2].Value.lpszW = body;
    properties[3].ulPropTag = PR_MESSAGE_FLAGS;
    properties[3].Value.l = MSGFLAG_UNSENT | MSGFLAG_READ;
    setMessageProperties(mapi, message.get(), 4, properties);
    return message;
}

static void saveDraft(const Mapi& mapi, IMAPIFolder* drafts, wchar_t** arguments) {
    auto message = createNote(mapi, drafts, arguments[0], arguments[1]);
    check(message->SaveChanges(KEEP_OPEN_READWRITE), "Save MAPI draft");
    SPropValue* entryRaw = nullptr;
    const auto entry = binaryProperty(mapi, message.get(), PR_ENTRYID, entryRaw);
    auto freeProps = [&mapi](SPropValue* values) { mapi.freeBuffer(values); };
    const std::unique_ptr<SPropValue, decltype(freeProps)> entryProperties(entryRaw, freeProps);
    printEntryId(entry);
    std::putchar('\n');
}

static void submitNote(const Mapi& mapi, IMAPISession* session, IMsgStore* store,
                       wchar_t** arguments) {
    auto* address = arguments[0];
    if (!*address || !std::wcschr(address, L'@') || std::wcspbrk(address, L" \t\r\n;,"))
        throw std::invalid_argument("mapi-send requires a single SMTP email address");
    auto freeProps = [&mapi](SPropValue* values) { mapi.freeBuffer(values); };
    SPropValue* outboxRaw = nullptr;
    const auto outboxId = binaryProperty(mapi, store, PR_IPM_OUTBOX_ENTRYID, outboxRaw);
    const std::unique_ptr<SPropValue, decltype(freeProps)> outboxProperties(outboxRaw, freeProps);
    auto outbox = openFolder(store, outboxId.cb, reinterpret_cast<LPENTRYID>(outboxId.lpb),
                            MAPI_MODIFY);
    auto message = createNote(mapi, outbox.get(), arguments[1], arguments[2]);

    SPropValue* sentRaw = nullptr;
    const auto sentId = binaryProperty(mapi, store, PR_IPM_SENTMAIL_ENTRYID, sentRaw);
    const std::unique_ptr<SPropValue, decltype(freeProps)> sentProperties(sentRaw, freeProps);
    SPropValue sentFolder{};
    sentFolder.ulPropTag = PR_SENTMAIL_ENTRYID;
    sentFolder.Value.bin = sentId;
    setMessageProperties(mapi, message.get(), 1, &sentFolder);

    IAddrBook* bookRaw = nullptr;
    check(session->OpenAddressBook(0, nullptr, AB_NO_DIALOG, &bookRaw), "Open MAPI address book");
    Interface<IAddrBook> book(bookRaw);
    wchar_t smtp[] = L"SMTP";
    ULONG idSize = 0;
    LPENTRYID idRaw = nullptr;
    check(book->CreateOneOff(reinterpret_cast<LPTSTR>(address), reinterpret_cast<LPTSTR>(smtp),
                            reinterpret_cast<LPTSTR>(address), MAPI_UNICODE | MAPI_SEND_NO_RICH_INFO,
                            &idSize, &idRaw), "Create MAPI SMTP recipient");
    auto freeId = [&mapi](ENTRYID* id) { mapi.freeBuffer(id); };
    const std::unique_ptr<ENTRYID, decltype(freeId)> entry(idRaw, freeId);

    SPropValue recipient[5]{};
    recipient[0].ulPropTag = PR_RECIPIENT_TYPE;
    recipient[0].Value.l = MAPI_TO;
    recipient[1].ulPropTag = PR_ENTRYID;
    recipient[1].Value.bin = {idSize, reinterpret_cast<LPBYTE>(entry.get())};
    recipient[2].ulPropTag = PR_DISPLAY_NAME_W;
    recipient[2].Value.lpszW = address;
    recipient[3].ulPropTag = PR_ADDRTYPE_W;
    recipient[3].Value.lpszW = smtp;
    recipient[4].ulPropTag = PR_EMAIL_ADDRESS_W;
    recipient[4].Value.lpszW = address;
    SizedADRLIST(1, recipients) = {1, {{0, 5, recipient}}};
    check(message->ModifyRecipients(MODRECIP_ADD, reinterpret_cast<LPADRLIST>(&recipients)),
          "Add MAPI recipient");
    check(message->SubmitMessage(0), "Submit MAPI message");
}

static unsigned hexDigit(wchar_t digit) {
    if (digit >= L'0' && digit <= L'9') return digit - L'0';
    if (digit >= L'A' && digit <= L'F') return digit - L'A' + 10;
    if (digit >= L'a' && digit <= L'f') return digit - L'a' + 10;
    throw std::runtime_error("An Outlook entry ID must contain only hexadecimal digits");
}

static void readBody(IMsgStore* store, const std::wstring& hex) {
    if (hex.empty() || hex.size() % 2)
        throw std::runtime_error("An Outlook entry ID must contain an even number of hexadecimal digits");
    std::vector<BYTE> id(hex.size() / 2);
    for (size_t i = 0; i < id.size(); ++i)
        id[i] = static_cast<BYTE>((hexDigit(hex[2*i]) << 4) | hexDigit(hex[2*i+1]));
    ULONG kind = 0;
    IUnknown* raw = nullptr;
    check(store->OpenEntry(static_cast<ULONG>(id.size()), reinterpret_cast<LPENTRYID>(id.data()),
                           nullptr, 0, &kind, &raw), "Open MAPI message");
    Interface<IUnknown> object(raw);
    if (kind != MAPI_MESSAGE || !object) throw std::runtime_error("The entry is not a MAPI message");
    auto* message = static_cast<IMessage*>(object.get());
    IStream* bodyRaw = nullptr;
    const HRESULT status = message->OpenProperty(PR_BODY_W, const_cast<LPIID>(&IID_IStream), 0, 0,
                                                reinterpret_cast<IUnknown**>(&bodyRaw));
    if (status == MAPI_E_NOT_FOUND) return;
    check(status, "Open MAPI message body");
    Interface<IStream> body(bodyRaw);
    std::wstring text;
    wchar_t buffer[2048];
    ULONG read = 0;
    do {
        check(body->Read(buffer, sizeof(buffer), &read), "Read MAPI message body");
        if (read % sizeof(wchar_t)) throw std::runtime_error("Outlook returned a malformed Unicode body");
        text.append(buffer, read / sizeof(wchar_t));
    } while (read != 0);
    if (!text.empty() && text.back() == L'\0') text.pop_back();
    output(text.data(), text.size());
}

static void mapiCommand(const std::wstring& command, wchar_t* profile, wchar_t** arguments) {
    Apartment apartment;
    Mapi mapi;
    if (command == L"mapi-check") {
        std::puts("Outlook Extended MAPI initialized");
        return;
    }
    IMAPISession* raw = nullptr;
    const ULONG flags = MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_UNICODE |
                        (profile ? 0 : MAPI_USE_DEFAULT);
    check(mapi.logon(0, reinterpret_cast<LPTSTR>(profile), nullptr, flags, &raw), "MAPILogonEx");
    Interface<IMAPISession> session(raw);
    IMAPITable* storesRaw = nullptr;
    HRESULT status = session->GetMsgStoresTable(MAPI_UNICODE, &storesRaw);
    if (status == MAPI_E_BAD_CHARWIDTH)
        status = session->GetMsgStoresTable(0, &storesRaw);
    check(status, "GetMsgStoresTable");
    Interface<IMAPITable> stores(storesRaw);
    if (command == L"mapi-stores") {
        printTable(mapi, stores.get(), PR_DISPLAY_NAME_W, ULONG_MAX);
        return;
    }
    SPropValue defaultValue{};
    defaultValue.ulPropTag = PR_DEFAULT_STORE;
    defaultValue.Value.b = TRUE;
    SRestriction restriction{};
    restriction.rt = RES_PROPERTY;
    restriction.res.resProperty = {RELOP_EQ, PR_DEFAULT_STORE, &defaultValue};
    check(stores->Restrict(&restriction, 0), "Find default MAPI store");
    SizedSPropTagArray(1, storeColumns) = {1, {PR_ENTRYID}};
    check(stores->SetColumns(reinterpret_cast<LPSPropTagArray>(&storeColumns), 0),
          "Select MAPI store entry ID");
    SRowSet* storeRowsRaw = nullptr;
    check(stores->QueryRows(1, 0, &storeRowsRaw), "Query default MAPI store");
    auto freeRows = [&mapi](SRowSet* rows) { mapi.freeRows(rows); };
    const std::unique_ptr<SRowSet, decltype(freeRows)> storeRows(storeRowsRaw, freeRows);
    if (!storeRows || !storeRows->cRows || !storeRows->aRow[0].cValues ||
        storeRows->aRow[0].lpProps[0].ulPropTag != PR_ENTRYID)
        throw std::runtime_error("The selected profile has no default MAPI store");
    const auto entry = storeRows->aRow[0].lpProps[0].Value.bin;
    IMsgStore* storeRaw = nullptr;
    const bool writing = command == L"mapi-draft";
    const bool sending = command == L"mapi-send";
    const auto storeFlags = MDB_NO_DIALOG | ((writing || sending) ? MDB_WRITE : 0);
    check(session->OpenMsgStore(0, entry.cb, reinterpret_cast<LPENTRYID>(entry.lpb),
                                nullptr, storeFlags, &storeRaw), "Open default MAPI store");
    Interface<IMsgStore> store(storeRaw);
    if (sending) {
        submitNote(mapi, session.get(), store.get(), arguments);
        return;
    }
    if (command == L"mapi-read") {
        readBody(store.get(), arguments[0]);
        return;
    }
    ULONG idSize = 0;
    LPENTRYID idRaw = nullptr;
    check(store->GetReceiveFolder(nullptr, 0, &idSize, &idRaw, nullptr), "GetReceiveFolder");
    auto freeId = [&mapi](ENTRYID* id) { mapi.freeBuffer(id); };
    const std::unique_ptr<ENTRYID, decltype(freeId)> id(idRaw, freeId);
    auto folder = openFolder(store.get(), idSize, id.get(), 0);
    const bool drafts = writing || command == L"mapi-drafts";
    if (drafts) {
        constexpr ULONG draftsEntryId = PROP_TAG(PT_BINARY, 0x36D7);
        SPropValue* rawProps = nullptr;
        const auto folderId = binaryProperty(mapi, folder.get(), draftsEntryId, rawProps);
        const auto freeProps = [&mapi](SPropValue* values) { mapi.freeBuffer(values); };
        const std::unique_ptr<SPropValue, decltype(freeProps)> props(rawProps, freeProps);
        folder = openFolder(store.get(), folderId.cb, reinterpret_cast<LPENTRYID>(folderId.lpb),
                            writing ? MAPI_MODIFY : 0);
    }
    if (writing) {
        saveDraft(mapi, folder.get(), arguments);
        return;
    }
    IMAPITable* tableRaw = nullptr;
    status = folder->GetContentsTable(MAPI_UNICODE, &tableRaw);
    if (status == MAPI_E_BAD_CHARWIDTH)
        status = folder->GetContentsTable(0, &tableRaw);
    check(status, "GetContentsTable");
    Interface<IMAPITable> table(tableRaw);
    SSortOrderSet sort{};
    sort.cSorts = 1;
    sort.aSort[0] = {drafts ? PR_LAST_MODIFICATION_TIME : PR_MESSAGE_DELIVERY_TIME, TABLE_SORT_DESCEND};
    check(table->SortTable(&sort, 0), "Sort MAPI inbox");
    printTable(mapi, table.get(), PR_SUBJECT_W, 50, true);
}

static bool validComCommand(const std::wstring& command, int arguments) {
    if (command == L"version" || command == L"folders" || command == L"inbox")
        return arguments == 0;
    if (command == L"exchange-status") return arguments == 1 || arguments == 2;
    if (command == L"sync") return arguments == 0 || arguments == 1;
    if (command == L"draft" || command == L"send") return arguments == 3;
    return command == L"com" && arguments >= 1;
}

static void comBatch(int argc, wchar_t** argv, bool batch) {
    std::vector<int> commands;
    for (int start = 0; start < argc;) {
        int end = start + 1;
        while (end < argc && (!batch || std::wcscmp(argv[end], L"--") != 0)) ++end;
        if (!validComCommand(argv[start], end - start - 1) || (end < argc && end + 1 == argc))
            throw std::invalid_argument("Invalid Outlook COM command or arguments");
        commands.push_back(start);
        argv[end] = nullptr;
        start = end + 1;
    }
    Apartment apartment;
    auto app = automation(L"Outlook.Application", CLSCTX_LOCAL_SERVER, "Outlook.Application", true);
    for (int start : commands)
        comCommand(app.get(), argv[start], argv + start + 1);
}

int wmain(int argc, wchar_t** argv) {
    try {
        if (_setmode(_fileno(stdout), _O_BINARY) == -1)
            throw std::runtime_error("Could not configure UTF-8 output");
        const std::wstring command = argc > 1 ? argv[1] : L"";
        if (command == L"mapi-send" && (argc == 5 || argc == 6)) {
            mapiCommand(command, argc == 6 ? argv[5] : nullptr, argv + 2);
        } else if (command == L"mapi-draft" && (argc == 4 || argc == 5)) {
            mapiCommand(command, argc == 5 ? argv[4] : nullptr, argv + 2);
        } else if (command == L"mapi-read" && (argc == 3 || argc == 4)) {
            mapiCommand(command, argc == 4 ? argv[3] : nullptr, argv + 2);
        } else if ((command == L"mapi-check" && argc == 2) ||
                   ((command == L"mapi-stores" || command == L"mapi-inbox" ||
                     command == L"mapi-drafts") && argc <= 3)) {
            mapiCommand(command, argc == 3 ? argv[2] : nullptr, nullptr);
        } else if (argc == 2 && command == L"com-check") {
            comCheck();
        } else if (argc >= 3 && command == L"batch") {
            comBatch(argc - 2, argv + 2, true);
        } else if (validComCommand(command, argc - 2)) {
            comBatch(argc - 1, argv + 1, false);
        } else {
            throw std::invalid_argument("Invalid command or arguments");
        }
    } catch (const std::invalid_argument& error) {
        std::fprintf(stderr, "outlook: %s. Run outlook --help.\n", error.what());
        return 2;
    } catch (const std::exception& error) {
        std::fprintf(stderr, "outlook: %s\n", error.what());
        return 1;
    }
}
