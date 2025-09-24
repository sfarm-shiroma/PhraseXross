# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Build configuration can be switched at build-time (Release by default)
ARG BUILD_CONFIGURATION=Release

# csproj を先にコピーして復元
COPY PhraseXross.csproj .
RUN dotnet restore

# 残りをコピーして publish
COPY . .
# Generate portable PDBs and symbols even for Release; use BUILD_CONFIGURATION for Debug when needed
RUN dotnet publish -c ${BUILD_CONFIGURATION} -o /app/publish /p:UseAppHost=false /p:DebugType=portable /p:DebugSymbols=true

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app

# コンテナ内は 8080 で待受
ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

# Allow diagnostics and install .NET debugger (vsdbg) for VS Code attach
ENV DOTNET_EnableDiagnostics=1
RUN apt-get update \
	&& apt-get install -y --no-install-recommends curl ca-certificates unzip \
	&& rm -rf /var/lib/apt/lists/* \
	&& curl -sSL https://aka.ms/getvsdbgsh | bash /dev/stdin -v latest -l /vsdbg

COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "PhraseXross.dll"]

# Optional: Log environment variables for debugging
RUN printenv > /app/env.debug