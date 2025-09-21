# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# csproj を先にコピーして復元
COPY PhraseXross.csproj .
RUN dotnet restore

# 残りをコピーして publish
COPY . .
RUN dotnet publish -c Release -o /app/publish /p:UseAppHost=false

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app

# コンテナ内は 8080 で待受
ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "PhraseXross.dll"]

# Optional: Log environment variables for debugging
RUN printenv > /app/env.debug