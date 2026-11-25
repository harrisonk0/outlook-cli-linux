CXX64 ?= x86_64-w64-mingw32-g++
CXX32 ?= i686-w64-mingw32-g++
FLAGS = -std=c++17 -Wall -Wextra -Werror -O2 -static -municode
LIBS = -lole32 -loleaut32 -luuid
all: outlook-x64.exe outlook-x86.exe
outlook-x64.exe: outlook.cpp
	$(CXX64) $(FLAGS) -o $@ $< $(LIBS)
outlook-x86.exe: outlook.cpp
	$(CXX32) $(FLAGS) -o $@ $< $(LIBS)

VERSION ?= 0.1.0
PACKAGE = outlook-cli-linux-v$(VERSION)

.PHONY: all dist
dist: all
	install -d dist/$(PACKAGE)
	install -m 755 outlook outlook-x64.exe outlook-x86.exe dist/$(PACKAGE)/
	install -m 644 README.md dist/$(PACKAGE)/
	tar -czf dist/$(PACKAGE).tar.gz -C dist $(PACKAGE)
	cd dist && sha256sum $(PACKAGE).tar.gz > SHA256SUMS
