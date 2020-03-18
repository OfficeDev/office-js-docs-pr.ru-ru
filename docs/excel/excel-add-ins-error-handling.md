---
title: Обработка ошибок
description: Изучите логику обработки ошибок API JavaScript для Excel, чтобы учитывать ошибки времени выполнения.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bee5824d8854a55d5ac4041be1335ce239b31a9e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717168"
---
# <a name="error-handling"></a><span data-ttu-id="a375b-103">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="a375b-103">Error handling</span></span>

<span data-ttu-id="a375b-p101">При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.</span><span class="sxs-lookup"><span data-stu-id="a375b-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="a375b-106">Для получения дополнительных сведений о `sync()` методе и асинхронной природе API JavaScript для Excel ознакомьтесь [с основными концепциями программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="a375b-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="a375b-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="a375b-107">Best practices</span></span>

<span data-ttu-id="a375b-p102">В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`. Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="a375b-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="a375b-110">Ошибки API</span><span class="sxs-lookup"><span data-stu-id="a375b-110">API errors</span></span>

<span data-ttu-id="a375b-111">Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства:</span><span class="sxs-lookup"><span data-stu-id="a375b-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="a375b-p103">**code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.</span><span class="sxs-lookup"><span data-stu-id="a375b-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="a375b-115">**message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке.</span><span class="sxs-lookup"><span data-stu-id="a375b-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="a375b-116">Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.</span><span class="sxs-lookup"><span data-stu-id="a375b-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="a375b-117">**debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.</span><span class="sxs-lookup"><span data-stu-id="a375b-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="a375b-118">Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере.</span><span class="sxs-lookup"><span data-stu-id="a375b-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="a375b-119">Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="a375b-119">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="a375b-120">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="a375b-120">Error Messages</span></span>

<span data-ttu-id="a375b-121">В таблице ниже перечислены ошибки, которые может возвращать API.</span><span class="sxs-lookup"><span data-stu-id="a375b-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="a375b-122">error.code</span><span class="sxs-lookup"><span data-stu-id="a375b-122">error.code</span></span> | <span data-ttu-id="a375b-123">error.message</span><span class="sxs-lookup"><span data-stu-id="a375b-123">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="a375b-124">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="a375b-124">InvalidArgument</span></span> |<span data-ttu-id="a375b-125">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="a375b-125">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="a375b-126">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="a375b-126">InvalidRequest</span></span>  |<span data-ttu-id="a375b-127">Не удается обработать запрос.</span><span class="sxs-lookup"><span data-stu-id="a375b-127">Cannot process the request.</span></span>|
|<span data-ttu-id="a375b-128">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="a375b-128">InvalidReference</span></span>|<span data-ttu-id="a375b-129">Эта ссылка недопустима для текущей операции.</span><span class="sxs-lookup"><span data-stu-id="a375b-129">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="a375b-130">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="a375b-130">InvalidBinding</span></span>  |<span data-ttu-id="a375b-131">Эта привязка объектов недопустима из-за предыдущих обновлений.</span><span class="sxs-lookup"><span data-stu-id="a375b-131">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="a375b-132">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="a375b-132">InvalidSelection</span></span>|<span data-ttu-id="a375b-133">Выбранный фрагмент недопустим для этой операции.</span><span class="sxs-lookup"><span data-stu-id="a375b-133">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="a375b-134">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="a375b-134">Unauthenticated</span></span> |<span data-ttu-id="a375b-135">Требуемые сведения о проверке подлинности отсутствуют или недопустимы.</span><span class="sxs-lookup"><span data-stu-id="a375b-135">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="a375b-136">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="a375b-136">AccessDenied</span></span> |<span data-ttu-id="a375b-137">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="a375b-137">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="a375b-138">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="a375b-138">ItemNotFound</span></span> |<span data-ttu-id="a375b-139">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="a375b-139">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="a375b-140">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="a375b-140">ActivityLimitReached</span></span>|<span data-ttu-id="a375b-141">Достигнут предел действий.</span><span class="sxs-lookup"><span data-stu-id="a375b-141">Activity limit has been reached.</span></span>|
|<span data-ttu-id="a375b-142">GeneralException</span><span class="sxs-lookup"><span data-stu-id="a375b-142">GeneralException</span></span>|<span data-ttu-id="a375b-143">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="a375b-143">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="a375b-144">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="a375b-144">NotImplemented</span></span>  |<span data-ttu-id="a375b-145">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="a375b-145">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="a375b-146">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="a375b-146">ServiceNotAvailable</span></span>|<span data-ttu-id="a375b-147">Служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="a375b-147">The service is unavailable.</span></span>|
|<span data-ttu-id="a375b-148">Conflict</span><span class="sxs-lookup"><span data-stu-id="a375b-148">Conflict</span></span>|<span data-ttu-id="a375b-149">Запрос не удалось обработать из-за конфликта.</span><span class="sxs-lookup"><span data-stu-id="a375b-149">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="a375b-150">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="a375b-150">ItemAlreadyExists</span></span>|<span data-ttu-id="a375b-151">Создаваемый ресурс уже существует.</span><span class="sxs-lookup"><span data-stu-id="a375b-151">The resource being created already exists.</span></span>|
|<span data-ttu-id="a375b-152">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="a375b-152">UnsupportedOperation</span></span>|<span data-ttu-id="a375b-153">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a375b-153">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="a375b-154">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="a375b-154">RequestAborted</span></span>|<span data-ttu-id="a375b-155">Запрос прерван во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="a375b-155">The request was aborted during run time.</span></span>|
|<span data-ttu-id="a375b-156">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="a375b-156">ApiNotAvailable</span></span>|<span data-ttu-id="a375b-157">Запрашиваемый интерфейс API недоступен.</span><span class="sxs-lookup"><span data-stu-id="a375b-157">The requested API is not available.</span></span>|
|<span data-ttu-id="a375b-158">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="a375b-158">InsertDeleteConflict</span></span>|<span data-ttu-id="a375b-159">Операция вставки или удаления привела к конфликту.</span><span class="sxs-lookup"><span data-stu-id="a375b-159">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="a375b-160">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="a375b-160">InvalidOperation</span></span>|<span data-ttu-id="a375b-161">Выполняемая операция недопустима для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="a375b-161">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="a375b-162">См. также</span><span class="sxs-lookup"><span data-stu-id="a375b-162">See also</span></span>

- [<span data-ttu-id="a375b-163">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a375b-163">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a375b-164">Объект OfficeExtension.Error (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="a375b-164">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
