---
title: Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: fec75531997a6b496b442b72d8139a9c98f37cd5
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270840"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="6bcd8-102">Добавление ссылок на библиотеку JavaScript API для Office из сети доставки содержимого (CDN)</span><span class="sxs-lookup"><span data-stu-id="6bcd8-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="6bcd8-103">Помимо действий, описанных в этой статье, если необходимо использовать TypeScript, для получения Intellisense потребуется запустить указанную ниже команду в системном приглашении с поддержкой Node (или в окне Git Bash) из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="6bcd8-103">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="6bcd8-104">У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).</span><span class="sxs-lookup"><span data-stu-id="6bcd8-104">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```bash
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="6bcd8-105">Библиотека [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js.</span><span class="sxs-lookup"><span data-stu-id="6bcd8-105">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="6bcd8-106">Простейший способ добавить ссылку на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:</span><span class="sxs-lookup"><span data-stu-id="6bcd8-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="6bcd8-p102">`/1/` перед `office.js` в URL-адресе CDN указывает, что необходимо использовать последний добавочный выпуск файла Office.js версии 1. Так как API JavaScript для Office обеспечивает обратную совместимость, в последнем выпуске будут и дальше поддерживаться элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, см. статью [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="6bcd8-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="6bcd8-p103">Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.</span><span class="sxs-lookup"><span data-stu-id="6bcd8-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="6bcd8-p104">Разрабатывая надстройку для ведущего приложения Office, ссылайтесь на API JavaScript для Office из раздела `<head>` страницы. Это гарантирует, что API полностью инициализируется раньше всех элементов основного текста. Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации. Если надстройка не активируется в течение этого срока, будет считаться, что она не отвечает, и пользователь увидит сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="6bcd8-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="6bcd8-116">См. также</span><span class="sxs-lookup"><span data-stu-id="6bcd8-116">See also</span></span>

- [<span data-ttu-id="6bcd8-117">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="6bcd8-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="6bcd8-118">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="6bcd8-118">JavaScript API for Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
    
