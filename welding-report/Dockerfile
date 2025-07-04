FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base

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
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["welding-report/welding-report.csproj", "welding-report/"]

RUN dotnet restore "./welding-report/welding-report.csproj"
COPY . .
WORKDIR "/src/welding-report"
RUN dotnet build "./welding-report.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./welding-report.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

# Стадия архивации
FROM publish AS archive
WORKDIR /app/publish
RUN tar -czf /app/welding.tar.gz .

FROM base AS final
WORKDIR /app
COPY --from=archive /app/welding.tar.gz /app/welding.tar.gz
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "welding-report.dll"]