Надстройки часто кэшируются в Office на Mac по соображениям производительности. Как правило, для очистки кэша необходимо перезагрузить надстройку. Если в одном документе несколько надстроек, автоматическая очистка кэша может не сработать при перезагрузке.

### <a name="use-the-personality-menu-to-clear-the-cache"></a>Очистка кэша с помощью меню личных данных

Вы можете очистить кэш с помощью меню личных данных любой надстройки области задач. Тем не менее, поскольку меню личных данных не поддерживается в надстройках Outlook, вы можете попробовать [очистить кэш вручную](#clear-the-cache-manually) , если вы используете Outlook.

- Откройте меню личных данных. Затем выберите **Очистить кэш веб-сайта**.
    > [!NOTE]
    > Чтобы открыть меню личных данных, необходимо запустить macOS версии 10.13.6 или более поздней.

    ![Снимок экрана: параметр "Очистить кэш веб-сайта" в меню личных данных.](../images/mac-clear-cache-menu.png)

### <a name="clear-the-cache-manually"></a>Очистка кэша вручную

Вы также можете очистить кэш вручную, удалив все содержимое папки `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. Найдите эту папку через терминал.

> [!NOTE]
> Если эта папка не существует, проверьте наличие следующих папок через терминал и, если она найдена, удалите содержимое папки.
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`, где `{host}` — это приложение Office (например, `Excel`)
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`, где `{host}` — это приложение Office (например, `Excel`)
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
>
> Чтобы найти эти папки с помощью Finder, необходимо задать параметр Finder для отображения скрытых файлов. Finder отображает папки в **каталоге Контейнеры** по названию продукта, например **Microsoft Excel** , а не **com.microsoft.Excel**.