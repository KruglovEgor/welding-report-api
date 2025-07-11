# Welding Report API

Веб-приложение для автоматической генерации отчетов по сварочным работам и заявкам на основе данных из системы Redmine.

## Описание

Приложение представляет собой REST API сервис, построенный на ASP.NET Core 8.0, который автоматизирует процесс создания различных типов отчетов:

- **Отчеты по заявкам** - генерация Word документов на основе данных из pr_redmine
- **Отчеты по сварочным работам** - генерация Excel файлов на основе данных из svarka
- **Отчеты СУПР** - генерация Excel файлов для системы СУПР

## Основные возможности

### 📄 Генерация отчетов
- **Word отчеты** - создание документов по заявкам с автоматическим заполнением данных
- **Excel отчеты** - создание табличных отчетов по сварочным работам и проектам
- **Автоматическая отправка** - возможность отправки отчетов на email

### 🔗 Интеграции
- **Redmine API** - получение данных из различных инстансов Redmine
- **Email сервис** - отправка отчетов через SMTP (Yandex)
- **Файловое хранилище** - временное хранение и кэширование файлов

### 🎨 Шаблоны
- Поддержка шаблонов Word (.docx) и Excel (.xlsx)
- Автоматическое заполнение данных из API
- Настройка стилей и форматирования

## Технологический стек

- **.NET 8.0** - основная платформа
- **ASP.NET Core** - веб-фреймворк
- **DocumentFormat.OpenXml** - работа с Word документами
- **EPPlus** - работа с Excel файлами
- **SkiaSharp** - обработка изображений
- **Swagger** - документация API
- **Docker** - контейнеризация

## Структура проекта

```
welding-report/
├── Controllers/           # API контроллеры
├── Models/               # Модели данных и настройки
├── Services/             # Бизнес-логика и сервисы
│   ├── Request/         # Сервисы для работы с заявками
│   ├── Welding/         # Сервисы для сварочных работ
│   └── Supr/            # Сервисы для СУПР
├── Resources/           # Шаблоны и ресурсы
│   └── Templates/       # Шаблоны отчетов
└── Program.cs           # Точка входа приложения
```

## API Endpoints

### Отчеты по заявкам
- `GET /api/WeldingReport/generate-issue-from-request` - генерация Word отчета по заявке

### Отчеты по сварочным работам
- `GET /api/WeldingReport/generate-issue-from-welding` - генерация Excel отчета по задаче
- `GET /api/WeldingReport/generate-project-from-welding` - генерация Excel отчета по проекту

### Отчеты СУПР
- `GET /api/WeldingReport/generate-group-from-supr` - генерация Excel отчета СУПР

## Конфигурация

Основные настройки находятся в `appsettings.json`:

- **AppSettings** - пути к файлам, настройки изображений
- **EmailSettings** - настройки SMTP сервера
- **RedmineSettings** - URL и API ключи для Redmine
- **SuprSignatures** - подписи для отчетов СУПР

## Запуск

### Локальная разработка
```bash
cd welding-report
dotnet run
```

### Docker
```bash
docker build -t welding-report .
docker run -p 8080:8080 welding-report
```

## Документация API

После запуска приложения документация Swagger доступна по адресу:
- http://localhost:8080/swagger

## Требования

- .NET 8.0 SDK
- Docker (опционально)
- Доступ к Redmine API
- Настроенный SMTP сервер для отправки email

## Лицензия

Проект использует EPPlus в некоммерческом режиме. 