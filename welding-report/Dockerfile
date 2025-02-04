FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER $APP_UID
WORKDIR /app
EXPOSE 8080
EXPOSE 8081

RUN mkdir -p /app/uploads && chmod 777 /app/uploads

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
ENTRYPOINT ["dotnet", "welding-report.dll"]