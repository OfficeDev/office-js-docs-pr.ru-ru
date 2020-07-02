---
title: Обработка ошибок
description: Изучите логику обработки ошибок API JavaScript для Excel, чтобы учитывать ошибки времени выполнения.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 8d410ae7eea315e14383b5aa08373ede3768cace
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006446"
---
# <a name="error-handling"></a><span data-ttu-id="ce6fc-103">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="ce6fc-103">Error handling</span></span>

<span data-ttu-id="ce6fc-104">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-104">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="ce6fc-105">Doing so is critical, due to the asynchronous nature of the API.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-105">Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="ce6fc-106">Для получения дополнительных сведений о `sync()` методе и асинхронной природе API JavaScript для Excel ознакомьтесь [с основными концепциями программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="ce6fc-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="ce6fc-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="ce6fc-107">Best practices</span></span>

<span data-ttu-id="ce6fc-108">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-108">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`.</span></span> <span data-ttu-id="ce6fc-109">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-109">We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="ce6fc-110">Ошибки API</span><span class="sxs-lookup"><span data-stu-id="ce6fc-110">API errors</span></span>

<span data-ttu-id="ce6fc-111">Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства:</span><span class="sxs-lookup"><span data-stu-id="ce6fc-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="ce6fc-112">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-112">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list.</span></span> <span data-ttu-id="ce6fc-113">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-113">For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation.</span></span> <span data-ttu-id="ce6fc-114">Error codes are not localized.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-114">Error codes are not localized.</span></span>

- <span data-ttu-id="ce6fc-115">**message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="ce6fc-116">Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="ce6fc-117">**debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="ce6fc-118">Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="ce6fc-119">Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-119">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="ce6fc-120">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="ce6fc-120">Error Messages</span></span>

<span data-ttu-id="ce6fc-121">В таблице ниже перечислены ошибки, которые может возвращать API.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="ce6fc-122">error.code</span><span class="sxs-lookup"><span data-stu-id="ce6fc-122">error.code</span></span> | <span data-ttu-id="ce6fc-123">error.message</span><span class="sxs-lookup"><span data-stu-id="ce6fc-123">error.message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="ce6fc-124">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="ce6fc-125">Достигнут предел действий.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="ce6fc-126">Запрашиваемый интерфейс API недоступен.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-126">The requested API is not available.</span></span>|
|`Conflict`|<span data-ttu-id="ce6fc-127">Запрос не удалось обработать из-за конфликта.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-127">Request could not be processed because of a conflict.</span></span>|
|`GeneralException`|<span data-ttu-id="ce6fc-128">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-128">There was an internal error while processing the request.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="ce6fc-129">Операция вставки или удаления привела к конфликту.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-129">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="ce6fc-130">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-130">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="ce6fc-131">Эта привязка объектов недопустима из-за предыдущих обновлений.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-131">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="ce6fc-132">Выполняемая операция недопустима для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-132">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="ce6fc-133">Эта ссылка недопустима для текущей операции.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-133">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="ce6fc-134">Не удается обработать запрос.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-134">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="ce6fc-135">Выбранный фрагмент недопустим для этой операции.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-135">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="ce6fc-136">Создаваемый ресурс уже существует.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-136">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="ce6fc-137">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-137">The requested resource doesn't exist.</span></span>|
|`NotImplemented`  |<span data-ttu-id="ce6fc-138">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-138">The requested feature isn't implemented.</span></span>|
|`RequestAborted`|<span data-ttu-id="ce6fc-139">Запрос прерван во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-139">The request was aborted during run time.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="ce6fc-140">Служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-140">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="ce6fc-141">Требуемые сведения о проверке подлинности отсутствуют или недопустимы.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-141">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="ce6fc-142">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-142">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="ce6fc-143">Этот тип листа не поддерживает эту операцию, так как он является макросом или листом диаграммы.</span><span class="sxs-lookup"><span data-stu-id="ce6fc-143">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="ce6fc-144">См. также</span><span class="sxs-lookup"><span data-stu-id="ce6fc-144">See also</span></span>

- [<span data-ttu-id="ce6fc-145">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ce6fc-145">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ce6fc-146">Объект OfficeExtension.Error (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="ce6fc-146">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview)
