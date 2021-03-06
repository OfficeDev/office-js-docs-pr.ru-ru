---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на библиотеку API JavaScript Office и определения типов в надстройки.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 346a34c0cbc31b5e569a5106dcd2bc01593b114a
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505194"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="2946a-103">Ссылки на библиотеку API JavaScript для Office </span><span class="sxs-lookup"><span data-stu-id="2946a-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="2946a-104">Библиотека [API JavaScript](../reference/javascript-api-for-office.md) Office предоставляет API, которые ваша надстройка может использовать для взаимодействия с приложением Office.</span><span class="sxs-lookup"><span data-stu-id="2946a-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="2946a-105">Самый простой способ ссылки на библиотеку — использовать сеть доставки контента (CDN), добавив следующий тег в разделе `<script>` `<head>` вашей HTML-страницы:</span><span class="sxs-lookup"><span data-stu-id="2946a-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="2946a-106">При первом загрузке надстройки будут загружаться и кэшируются файлы API Office JavaScript, чтобы убедиться, что она использует самые последние реализации Office.js и связанных с ними файлов для указанной версии.</span><span class="sxs-lookup"><span data-stu-id="2946a-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2946a-107">Чтобы убедиться, что API полностью инициализирован перед любыми элементами тела, необходимо ссылаться на API JavaScript Office из раздела `<head>` страницы.</span><span class="sxs-lookup"><span data-stu-id="2946a-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="2946a-108">Версия API и обратная совместимость</span><span class="sxs-lookup"><span data-stu-id="2946a-108">API versioning and backward compatibility</span></span>

<span data-ttu-id="2946a-109">В предыдущем фрагменте HTML перед URL-адресом CDN указывается последний дополнительный выпуск в версии `/1/` `office.js` 1 Office.js.</span><span class="sxs-lookup"><span data-stu-id="2946a-109">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="2946a-110">Так как API JavaScript Office поддерживает обратную совместимость, последний выпуск будет по-прежнему поддерживать участников API, которые были представлены ранее в версии 1.</span><span class="sxs-lookup"><span data-stu-id="2946a-110">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="2946a-111">Если необходимо обновить существующий проект, см. в статью Обновление версии [API JavaScript Office и файлы схемы манифеста.](update-your-javascript-api-for-office-and-manifest-schema-version.md)</span><span class="sxs-lookup"><span data-stu-id="2946a-111">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="2946a-p103">Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.</span><span class="sxs-lookup"><span data-stu-id="2946a-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="2946a-114">Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="2946a-114">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="2946a-115">Включение IntelliSense для проекта TypeScript</span><span class="sxs-lookup"><span data-stu-id="2946a-115">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="2946a-116">Помимо ссылок на API JavaScript Office, как описано выше, вы также можете включить IntelliSense для проекта надстройки TypeScript с помощью определений типа [из DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="2946a-116">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="2946a-117">Для этого запустите следующую команду в системном запросе с поддержкой узла (или в окне баш git) из корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="2946a-117">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="2946a-118">У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).</span><span class="sxs-lookup"><span data-stu-id="2946a-118">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="2946a-119">Предварительные API</span><span class="sxs-lookup"><span data-stu-id="2946a-119">Preview APIs</span></span>

<span data-ttu-id="2946a-120">Новые API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и необходимости отзыва пользователей.</span><span class="sxs-lookup"><span data-stu-id="2946a-120">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="2946a-121">См. также</span><span class="sxs-lookup"><span data-stu-id="2946a-121">See also</span></span>

- [<span data-ttu-id="2946a-122">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2946a-122">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="2946a-123">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2946a-123">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
