# Dükkan Defter — çok aşamalı derleme
# Derleme: docker build -t dukkan-defter:latest .
# Çalıştırma: docker compose up -d (veya aşağıdaki örnek run komutu)

FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

COPY ["DukkanDefterOCR.csproj", "./"]
RUN dotnet restore "DukkanDefterOCR.csproj"

COPY . .
RUN dotnet publish "DukkanDefterOCR.csproj" -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS final
WORKDIR /app

# Kalıcı SQLite için (volume ile eşleştirin)
RUN mkdir -p /app/data

ENV ASPNETCORE_URLS=http://+:8080
ENV ASPNETCORE_ENVIRONMENT=Production
# Nginx TLS sonlandırma arkasında redirect döngüsünü önlemek için
ENV BEHIND_REVERSE_PROXY=true
# Örnek: docker run -e ConnectionStrings__DefaultConnection="Data Source=/app/data/app.db" ...

EXPOSE 8080

COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "DukkanDefterOCR.dll"]
