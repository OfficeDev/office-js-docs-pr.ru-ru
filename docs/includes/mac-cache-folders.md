Надстройки часто кэшются в Office Mac по соображениям производительности. Как правило, для очистки кэша необходимо перезагрузить надстройку. Если в одном документе несколько надстроек, автоматическая очистка кэша может не сработать при перезагрузке.

### <a name="use-the-personality-menu-to-clear-the-cache"></a>Чтобы очистить кэш, используйте меню личности

Вы можете очистить кэш с помощью меню личных данных любой надстройки области задач. Однако, так как меню личности не поддерживается в Outlook надстройки, вы можете попробовать параметр, чтобы очистить кэш вручную, если вы используете Outlook.[](#clear-the-cache-manually)

- Откройте меню личных данных. Затем выберите **Очистить кэш веб-сайта**.
    > [!NOTE]
    > Меню личных данных доступно в macOS версии 10.13.6 и более поздних версиях.

    ![Снимок экрана: параметр "Очистить кэш веб-сайта" в меню личных данных.](../images/mac-clear-cache-menu.png)

### <a name="clear-the-cache-manually"></a>Очистка кэша вручную

Вы также можете очистить кэш вручную, удалив все содержимое папки `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. И посмотрите на эту папку с помощью терминала.

> [!NOTE]
> Если эта папка не существует, проверьте следующие папки с помощью терминала и, если она найдена, удалите содержимое папки.
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`, где `{host}` — это приложение Office (например, `Excel`)
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`, где `{host}` — это приложение Office (например, `Excel`)
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
>
> Чтобы найти эти папки с помощью Finder, необходимо настроить Finder для показа скрытых файлов. Finder отображает папки в каталоге **Контейнеры** по имени продукта, например Microsoft Excel вместо **com.microsoft.Excel**.