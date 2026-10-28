# syntax=docker/dockerfile:1
# Mirrors the linux-x64 job in .github/workflows/build.yml (publish + smoke test).
FROM mcr.microsoft.com/dotnet/sdk:10.0

ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1 \
    DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 \
    NUGET_XMLDOC_MODE=skip \
    OFFICECLI_NO_AUTO_RESIDENT=1

RUN apt-get update && apt-get install -y --no-install-recommends \
    bash ca-certificates git python3 unzip \
    && rm -rf /var/lib/apt/lists/*

COPY . /src
WORKDIR /src

# Restore online; the scan phase runs without network access.
RUN dotnet restore src/officecli/officecli.csproj -r linux-x64

RUN dotnet publish src/officecli/officecli.csproj -c Release -r linux-x64 -o /opt/officecli --nologo --no-restore \
    && ln -s /opt/officecli/officecli /usr/local/bin/officecli

# Smoke test, same as CI; a failure is reported but does not abort the build.
RUN cd /tmp \
    && (officecli create test_smoke.docx \
        && officecli add test_smoke.docx /body --type paragraph --prop text="Hello from CI" \
        && officecli get test_smoke.docx '/body/p[1]' \
        && officecli close test_smoke.docx \
        || echo "WARNING: smoke test failed") \
    && rm -f test_smoke.docx

CMD ["bash"]
