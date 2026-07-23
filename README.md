# ActGeneratorPC Portable for Windows

Portable-версия генератора актов на PySide6 с автоматическим обновлением
через GitHub Releases. Установка не нужна: распакуйте ZIP в любую доступную
для записи папку и запустите `ActGeneratorPC.exe`.

## Что обновляется

При обновлении файлы из нового `ActGeneratorPC-portable.zip` накладываются на
текущую portable-папку:

- `ActGeneratorPC.exe` и папка `_internal`;
- `Data` со списками и иконками;
- `Templates` с шаблонами Word;
- конфигурация и остальные файлы релиза.

`Act_Ready` намеренно исключена: в ней находятся уже созданные пользователем
документы. Лишние локальные файлы не удаляются.

Перед установкой архив обязательно проверяется по
`ActGeneratorPC-portable.zip.sha256`. Пути внутри ZIP также проверяются перед
распаковкой.

## Подготовка проекта

1. Положите исходные списки и иконки в `Data`.
2. Положите шаблоны `ABP_MKTF.docx` и `Validator_MID.docx` в `Templates`.
3. Создайте публичный репозиторий GitHub и загрузите туда весь проект.
4. Для локальной сборки укажите `OWNER/REPOSITORY` в
   `update_config.json`. GitHub Actions заполняет это поле автоматически.

Автообновление без дополнительной авторизации рассчитано на публичный
репозиторий.

## Локальный запуск

```powershell
python -m pip install -r requirements.txt
.\run-dev.ps1
```

## Локальная portable-сборка

```powershell
.\build.ps1
```

Результат:

- `dist\ActGeneratorPC-portable.zip`;
- `dist\ActGeneratorPC-portable.zip.sha256`.

Если зависимости уже установлены:

```powershell
.\build.ps1 -SkipInstall
```

## Выпуск новой версии

Версия приложения задаётся в `version.py`. При выпуске через GitHub Actions
она автоматически берётся из имени тега.

```powershell
git tag v1.0.0
git push origin v1.0.0
```

Workflow на Windows:

1. подставит `1.0.0` в `version.py`;
2. запишет текущий `OWNER/REPOSITORY` в конфигурацию;
3. соберёт portable-папку и ZIP;
4. создаст SHA-256;
5. опубликует оба файла в GitHub Release.

При следующем запуске установленная portable-версия проверит `latest`
release, предложит обновление, завершится, обновит весь состав релиза и
запустится снова.

## Структура

```text
ActGeneratorPC-portable-source/
├── app.py
├── update_manager.py
├── version.py
├── update_config.json
├── ActGeneratorPC.spec
├── build.ps1
├── Data/
├── Templates/
├── Act_Ready/
└── .github/workflows/release.yml
```
