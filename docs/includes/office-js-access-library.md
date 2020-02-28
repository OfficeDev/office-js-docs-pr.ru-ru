Доступ к библиотеке API JavaScript для Office можно получить через сеть доставки содержимого (CDN) Office JS по адресу `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. Чтобы использовать API JavaScript для Office на любой из веб-страниц надстройки, требуется указать ссылку на CDN в теге `<script>` тега `<head>` страницы.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

Дополнительные сведения о доступе к библиотеке API JavaScript для Office, в том числе о способе получения функции IntelliSense, см. в статье [Добавление ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).