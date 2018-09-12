---
title: Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 9d3328ba09e2f69e76bd55f21064d52a8537cfa9
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943903"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="5117c-102">Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)</span><span class="sxs-lookup"><span data-stu-id="5117c-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>


<span data-ttu-id="5117c-103">Библиотека [API JavaScript для Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js.</span><span class="sxs-lookup"><span data-stu-id="5117c-103">The [JavaScript API for Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="5117c-104">Простейший способ добавить ссылку на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:</span><span class="sxs-lookup"><span data-stu-id="5117c-104">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="5117c-p101">`/1/` перед `office.js` в URL-адресе CDN указывает, что необходимо использовать последний добавочный выпуск файла Office.js версии 1. Так как API JavaScript для Office обеспечивает обратную совместимость, в последнем выпуске будут и дальше поддерживаться элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, см. статью [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="5117c-p101">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="5117c-p102">Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.</span><span class="sxs-lookup"><span data-stu-id="5117c-p102">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="5117c-p103">Разрабатывая надстройку для ведущего приложения Office, ссылайтесь на API JavaScript для Office из раздела `<head>` страницы. Это гарантирует, что API полностью инициализируется раньше всех элементов основного текста. Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации. Если надстройка не активируется в течение этого срока, будет считаться, что она не отвечает, и пользователь увидит сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="5117c-p103">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="5117c-114">См. также</span><span class="sxs-lookup"><span data-stu-id="5117c-114">See also</span></span>

- [<span data-ttu-id="5117c-115">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="5117c-115">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="5117c-116">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="5117c-116">JavaScript API for Office</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
