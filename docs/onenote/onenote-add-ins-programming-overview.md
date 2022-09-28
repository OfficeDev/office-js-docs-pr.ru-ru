---
title: Обзор создания кода с помощью API JavaScript для OneNote
description: Узнайте об API OneNote JavaScript для надстроек OneNote в Интернете.
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: d44a01cf0f676057ca072cff74e2e80057f645f4
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092912"
---
# <a name="onenote-javascript-api-programming-overview"></a>Обзор создания кода с помощью API JavaScript для OneNote

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Компоненты надстройки Office

Надстройки состоят из двух указанных ниже основных компонентов.

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a browser control or iframe.

- An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.

### <a name="office-add-in--manifest--webpage"></a>Надстройка Office = манифест + веб-страница

![Надстройка Office состоит из манифеста и веб-страницы.](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Использование API JavaScript

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- **API для определенных клиентских приложений** для связанных с OneNote операций, доступ к которому осуществляется с помощью объекта `Application`.
- **Общий API**, используемый приложениями Office, доступ к которому осуществляется с помощью объекта `Document`.

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>Доступ к API для определенных клиентских приложений с помощью объекта *Application*

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. Получение экземпляра приложения из контекста.

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > Вызовы, которые методы совершают к API (например, `context.application.getActiveSection().pages;`), также добавляются в очередь.

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

Например,

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

Дополнительные сведения о `load`/`sync`шаблонах и других распространенных практиках, используемых в API JavaScript для OneNote, см. в статье [Использование модели API, зависящей от приложений](../develop/application-specific-api-model.md).

Сведения о поддерживаемых объектах и операциях OneNote см. в [справочнике по API](../reference/overview/onenote-add-ins-javascript-reference.md).

#### <a name="onenote-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для OneNote

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### <a name="accessing-the-common-api-through-the-document-object"></a>Получение доступа к общему API с помощью объекта *Document*

Для доступа к общему API, например к методам [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) и [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)), используйте объект `Document`.

Например:  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

Надстройки OneNote поддерживают только указанные ниже общие API.

| API | Примечания |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | Только `Office.CoercionType.Text` и `Office.CoercionType.Matrix` |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | Только `Office.CoercionType.Text`, `Office.CoercionType.Image` и `Office.CoercionType.Html` |
| [const mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | Параметры поддерживаются только контентными надстройками |
| [Office.context.document.settings.set(имя, значение);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | Параметры поддерживаются только контентными надстройками |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*Ни один.*|

Обычно общий API следует использовать, когда необходимые возможности не поддерживаются в API для определенных клиентских приложений. Дополнительные сведения об использовании общего API см. в статье [Общая объектная модель API JavaScript](../develop/office-javascript-api-object-model.md).

<a name="om-diagram"></a>

## <a name="onenote-object-model-diagram"></a>Схема объектной модели OneNote

На схеме ниже показаны возможности, которые на данный момент доступны в API JavaScript для OneNote .

  ![Схема объектной модели OneNote.](../images/onenote-om.png)

## <a name="see-also"></a>См. также

- [Разработка надстроек Office](../develop/develop-overview.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Создание первой надстройки OneNote](../quickstarts/onenote-quickstart.md)
- [Справочник по API JavaScript для OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
