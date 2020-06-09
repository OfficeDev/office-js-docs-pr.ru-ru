---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на библиотеку API JavaScript для Office и определение типов в надстройке.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 8bd011c140ce61581ad4b1d06a43b04ad437f5c7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609389"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="7c13e-103">Ссылки на библиотеку API JavaScript для Office </span><span class="sxs-lookup"><span data-stu-id="7c13e-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="7c13e-104">Библиотека [API JavaScript для Office](../reference/javascript-api-for-office.md) предоставляет API, которые надстройка может использовать для взаимодействия с ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="7c13e-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host.</span></span> <span data-ttu-id="7c13e-105">Самый простой способ добавить ссылку на библиотеку — использовать сеть доставки содержимого (CDN), добавив следующий `<script>` тег в `<head>` раздел страницы HTML:</span><span class="sxs-lookup"><span data-stu-id="7c13e-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="7c13e-106">Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первом запуске надстройки, чтобы убедиться в том, что используется самая последняя реализация Office. js и связанных с ней файлов для указанной версии.</span><span class="sxs-lookup"><span data-stu-id="7c13e-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7c13e-107">Необходимо ссылаться на API JavaScript для Office из `<head>` раздела страницы, чтобы убедиться, что API полностью инициализирован до элементов основного текста.</span><span class="sxs-lookup"><span data-stu-id="7c13e-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="7c13e-108">Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации.</span><span class="sxs-lookup"><span data-stu-id="7c13e-108">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="7c13e-109">Если надстройка не активируется в этом пороговом значении, она будет объявлена без ответа, а пользователю будет выведено сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="7c13e-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="7c13e-110">Управление версиями и обратная совместимость API</span><span class="sxs-lookup"><span data-stu-id="7c13e-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="7c13e-111">В предыдущем фрагменте кода HTML ( `/1/` перед в `office.js` URL-адресе CDN) указывает последний добавочный выпуск в версии 1 файла Office. js.</span><span class="sxs-lookup"><span data-stu-id="7c13e-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="7c13e-112">Так как API JavaScript для Office поддерживает обратную совместимость, последний выпуск по-прежнему будет поддерживать элементы API, представленные ранее в версии 1.</span><span class="sxs-lookup"><span data-stu-id="7c13e-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="7c13e-113">Если вам нужно обновить существующий проект, ознакомьтесь со статьей [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="7c13e-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="7c13e-p104">Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.</span><span class="sxs-lookup"><span data-stu-id="7c13e-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="7c13e-116">Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="7c13e-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="7c13e-117">Включение IntelliSense для проекта TypeScript</span><span class="sxs-lookup"><span data-stu-id="7c13e-117">Enabling Intellisense for a TypeScript project</span></span>

<span data-ttu-id="7c13e-118">Кроме ссылки на API JavaScript для Office, как описано выше, можно также включить функцию IntelliSense для проекта надстройки TypeScript, используя определения типов из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="7c13e-118">In addition to referencing the Office JavaScript API as described previously, you can also enable Intellisense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="7c13e-119">Для этого выполните следующую команду в командной строке с поддержкой узлов (или в окне Bash Git) из корневого каталога папки проекта.</span><span class="sxs-lookup"><span data-stu-id="7c13e-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="7c13e-120">У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).</span><span class="sxs-lookup"><span data-stu-id="7c13e-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> <span data-ttu-id="7c13e-121">Чтобы включить IntelliSense для предварительной версии API, используйте следующие команды в корневой папке проекта, [выполнив следующую](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) команду:</span><span class="sxs-lookup"><span data-stu-id="7c13e-121">To enable Intellisense for preview APIs, use the preview type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) by running the following command in the root of your project folder:</span></span> 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a><span data-ttu-id="7c13e-122">См. также</span><span class="sxs-lookup"><span data-stu-id="7c13e-122">See also</span></span>

- [<span data-ttu-id="7c13e-123">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="7c13e-123">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="7c13e-124">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="7c13e-124">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
