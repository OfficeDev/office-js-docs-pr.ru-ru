Если проект основан на node.js (то есть не разработан с сервером Visual Studio и сервером интернет-информации (IIS)), вы можете заставить Office на Windows использовать Edge Legacy или Internet Explorer для запуска надстройок, даже если у вас есть сочетание версий Windows и Office, которые обычно используют более недавний браузер. Дополнительные сведения о том, какие браузеры используются различными сочетаниями версий Windows и Office, см. в браузерах, используемых Office [надстройки.](../concepts/browsers-used-by-office-web-add-ins.md)

1. Если проект не *был* создан с помощью средства Yo Office, необходимо установить средство настройки office-addin-dev. Запустите следующую команду в командной подсказке.

    ```command&nbsp;line
    npm install office-addin-dev-settings --save-dev
    ```

    [!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

1. Укажите браузер, который Office использовать со следующей командой в командной подсказке в корне проекта. Замените относительный путь, который является только имям файла манифеста, если он находится в `<path-to-manifest>` корне проекта. Замените `<webview>` либо на , либо на `ie` `edge-legacy` .

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> <webview>
    ```

    Ниже приведен пример.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

    В командной строке должно быть видно сообщение о том, что тип веб-просмотров теперь за набором IE (или Edge Legacy).

1. По завершению установите Office с помощью браузера по умолчанию для сочетания версий Windows и Office со следующей командой.

    ```command&nbsp;line
    npx office-addin-dev-settings webview <path-to-manifest> default
    ```
