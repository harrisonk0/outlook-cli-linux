FROM debian:trixie-slim AS build
RUN --mount=type=bind,target=/src \
    apt-get update && apt-get install -y --no-install-recommends make g++-mingw-w64-x86-64 && \
    mkdir /build && cp /src/Makefile /src/outlook.cpp /build/ && \
    make -C /build outlook-x64.exe && rm -rf /var/lib/apt/lists/*

FROM debian:trixie-slim
ARG SODA_VERSION=soda-11.0-38-experimental
ARG SODA_SHA256=2852e5c53db23fba5bc3208f2115fc37d7ce09945154e400454f5b193b8ef644
RUN --mount=type=bind,target=/src \
    --mount=type=bind,from=build,source=/build,target=/build \
    --mount=type=secret,id=proxy-ca \
    set -eu; \
    apt-get update; apt-get install -y --no-install-recommends \
      bash ca-certificates curl xz-utils cabextract openssl tini util-linux \
      xvfb x11-utils xauth x11vnc novnc websockify openbox xdg-utils firefox-esr \
      fonts-liberation fonts-dejavu-core libunwind8 libgl1 libgl1-mesa-dri \
      libgnutls30t64 libfreetype6 libfontconfig1 libxrender1 libxrandr2 libxi6 \
      libxcursor1 libxinerama1 libxxf86vm1 libxcomposite1 libxfixes3 libvulkan1 \
      libasound2t64 libpulse0 libcups2t64 libdbus-1-3 libudev1 libusb-1.0-0; \
    rm -rf /var/lib/apt/lists/*; \
    if [ -s /run/secrets/proxy-ca ]; then \
      cat /etc/ssl/certs/ca-certificates.crt /run/secrets/proxy-ca > /tmp/build-ca.crt; \
      export CURL_CA_BUNDLE=/tmp/build-ca.crt; \
    fi; \
    mkdir /opt/soda /opt/outlook; \
    curl -fsSL --retry 3 "https://github.com/bottlesdevs/wine/releases/download/${SODA_VERSION}/${SODA_VERSION}-x86_64.tar.xz" -o /tmp/soda.tar.xz; \
    printf '%s  %s\n' "$SODA_SHA256" /tmp/soda.tar.xz | sha256sum -c -; \
    tar -xf /tmp/soda.tar.xz -C /opt/soda --strip-components=1; \
    rm -f /tmp/soda.tar.xz /tmp/build-ca.crt; \
    groupadd -g 10000 outlook; useradd -m -u 10000 -g outlook outlook; \
    install -d -o outlook -g outlook -m 0700 /state /run/outlook; \
    install -m 0755 /build/outlook-x64.exe /src/outlook /opt/outlook/; \
    install -m 0755 /src/docker/* /usr/local/bin/
ENV PATH=/opt/outlook:$PATH WINEPREFIX=/state/prefix OUTLOOK_WINE=/opt/soda/bin/wine OUTLOOK_ARCH=x64 \
    WINEDEBUG=-all DISPLAY=:88 XAUTHORITY=/run/outlook/Xauthority \
    XDG_RUNTIME_DIR=/run/outlook BROWSER=firefox-esr
USER outlook
WORKDIR /home/outlook
ENTRYPOINT ["/usr/bin/tini","-g","--","/usr/local/bin/entrypoint"]
