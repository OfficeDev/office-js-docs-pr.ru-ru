<span data-ttu-id="cda2b-101">Доступ к библиотеке API JavaScript для Office можно получить через сеть доставки содержимого (CDN) Office JS по адресу `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span><span class="sxs-lookup"><span data-stu-id="cda2b-101">The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="cda2b-102">Чтобы использовать API JavaScript для Office на любой из веб-страниц надстройки, требуется указать ссылку на CDN в теге `<script>` тега `<head>` страницы.</span><span class="sxs-lookup"><span data-stu-id="cda2b-102">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="cda2b-103">Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="cda2b-103">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

<span data-ttu-id="cda2b-104">Дополнительные сведения о доступе к библиотеке API JavaScript для Office, в том числе о способе получения функции IntelliSense, см. в статье [Добавление ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="cda2b-104">For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>