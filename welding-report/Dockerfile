FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 8080
EXPOSE 8081

# Устанавливаем зависимости для SkiaSharp и System.Drawing
RUN apt-get update && apt-get install -y \
    libfontconfig1 \
    libfreetype6 \
    libpng-dev \
    libjpeg-dev \
    libglib2.0-dev \
    libcairo2-dev \
    && rm -rf /var/lib/apt/lists/*

RUN mkdir -p /app/uploads /app/GeneratedReports && chmod -R 777 /app/uploads /app/GeneratedReports

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["welding-report.csproj", "."]
RUN dotnet restore "welding-report.csproj"
COPY . .
RUN dotnet build "welding-report.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "welding-report.csproj" -c Release -o /app/publish

# Финальный образ
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENV ASPNETCORE_URLS=http://+:8080;https://+:8081
VOLUME /app/uploads /app/GeneratedReports
ENTRYPOINT ["dotnet", "welding-report.dll"]