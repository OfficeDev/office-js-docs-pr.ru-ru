---
title: Обработка ошибок
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 23a70b1d66befb971c3c1394eb9162c19f2ee176
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348088"
---
# <a name="error-handling"></a><span data-ttu-id="64a8f-102">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="64a8f-102">Error handling</span></span>

<span data-ttu-id="64a8f-103">При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="64a8f-103">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="64a8f-104">Это очень важно из-за асинхронного характера API.</span><span class="sxs-lookup"><span data-stu-id="64a8f-104">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="64a8f-105">Дополнительные сведения о методе **sync()** и асинхронном характере API JavaScript для Excel см. в статье [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="64a8f-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="64a8f-106">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="64a8f-106">Best practices</span></span>

<span data-ttu-id="64a8f-107">В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="64a8f-107">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="64a8f-108">Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="64a8f-108">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

```js
Excel.run(function (context) { 
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);     
```

## <a name="api-errors"></a><span data-ttu-id="64a8f-109">Ошибки API</span><span class="sxs-lookup"><span data-stu-id="64a8f-109">API errors</span></span> 

<span data-ttu-id="64a8f-110">Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства.</span><span class="sxs-lookup"><span data-stu-id="64a8f-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span> 

- <span data-ttu-id="64a8f-111">**code**. Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes`или`Excel.ErrorCodes`.</span><span class="sxs-lookup"><span data-stu-id="64a8f-111">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="64a8f-112">Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции.</span><span class="sxs-lookup"><span data-stu-id="64a8f-112">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="64a8f-113">Коды ошибок не локализованы.</span><span class="sxs-lookup"><span data-stu-id="64a8f-113">Error codes are not localized.</span></span> 

- <span data-ttu-id="64a8f-114">**message**. Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке.</span><span class="sxs-lookup"><span data-stu-id="64a8f-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="64a8f-115">Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.</span><span class="sxs-lookup"><span data-stu-id="64a8f-115">The error message is not intended for end-user consumption; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end-users.</span></span>

- <span data-ttu-id="64a8f-116">**debugInfo**. Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять первопричину ошибки.</span><span class="sxs-lookup"><span data-stu-id="64a8f-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span> 

> [!NOTE]
> <span data-ttu-id="64a8f-117">Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере.</span><span class="sxs-lookup"><span data-stu-id="64a8f-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="64a8f-118">Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="64a8f-118">End-users will not see those error messages in the add-in taskpane or anywhere in the host application.</span></span>

## <a name="see-also"></a><span data-ttu-id="64a8f-119">См. также</span><span class="sxs-lookup"><span data-stu-id="64a8f-119">See also</span></span>

- [<span data-ttu-id="64a8f-120">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="64a8f-120">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="64a8f-121">Объект OfficeExtension.Error (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="64a8f-121">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/office/officeextension.error?view=office-js)
