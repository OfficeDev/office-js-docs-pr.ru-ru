Если проект основан на node.js (то есть не разработан с помощью Visual Studio и сервера IIS), вы можете принудительно использовать устаревшую версию Microsoft Edge или Internet Explorer для работы надстроек, даже если у вас есть сочетание версий Windows и Office, которые обычно используют более поздний браузер. Дополнительные сведения о том, какие браузеры используются различными сочетаниями версий Windows и Office, см. в разделе "Браузеры" [надстроек Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> Средство, которое используется для принудительного изменения в браузере, поддерживается только в канале бета-подписки Microsoft 365. Присоединяйтесь [к программе предварительной](https://insider.office.com/join/windows) оценки Office и выберите параметр **бета-канала** , чтобы получить доступ к бета-сборкам Office. См. [также о Office: какую версию Office я могу использовать?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> Строго, для этого средства `webview` (см. шаг **2**) требуется бета-канал. Средство имеет другие параметры, которые не имеют этого требования.

1. Если проект *не был создан* с помощью генератора [Yeoman для надстроек Office](../develop/yeoman-generator-overview.md) , необходимо установить средство office-addin-dev-settings. Выполните следующую команду в командной строке.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Укажите браузер, который будет использовать Office, с помощью следующей команды в командной строке в корневом каталоге проекта. Замените `<path-to-manifest>` относительный путь, который является только именем файла манифеста, если он находится в корне проекта. Замените `<webview>` на любой или `ie` `edge-legacy`.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Ниже приведен пример.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    В командной строке должно появиться сообщение о том, что для типа веб-представления теперь задано значение IE (или Устаревшая версия Edge).

1. Когда все будет готово, задайте office для возобновления работы с помощью браузера по умолчанию для сочетания версий Windows и Office с помощью следующей команды.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
