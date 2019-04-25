---
title: Обработка ошибок
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 87401773ad4a27bf0a30bc80b229d2879dd5234f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448163"
---
# <a name="error-handling"></a><span data-ttu-id="8205e-102">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="8205e-102">Error handling</span></span>

<span data-ttu-id="8205e-p101">При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.</span><span class="sxs-lookup"><span data-stu-id="8205e-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="8205e-105">Дополнительные сведения о методе **sync()** и асинхронном характере API JavaScript для Excel см. в статье [Основные понятия программирования с использованием API JavaScript для Excel](excel-add-ins-core-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="8205e-105">For more information about the **sync()** method and the asynchronous nature of Excel JavaScript API, see [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="8205e-106">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="8205e-106">Best practices</span></span>

<span data-ttu-id="8205e-p102">В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`. Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="8205e-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="8205e-109">Ошибки API</span><span class="sxs-lookup"><span data-stu-id="8205e-109">API errors</span></span>

<span data-ttu-id="8205e-110">Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства:</span><span class="sxs-lookup"><span data-stu-id="8205e-110">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="8205e-p103">**code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.</span><span class="sxs-lookup"><span data-stu-id="8205e-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="8205e-114">**message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке.</span><span class="sxs-lookup"><span data-stu-id="8205e-114">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="8205e-115">Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.</span><span class="sxs-lookup"><span data-stu-id="8205e-115">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="8205e-116">**debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.</span><span class="sxs-lookup"><span data-stu-id="8205e-116">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="8205e-117">Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере.</span><span class="sxs-lookup"><span data-stu-id="8205e-117">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="8205e-118">Эти сообщения об ошибках не будут отображаться для пользователей в области задач надстройки или в другом месте ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="8205e-118">End users will not see those error messages in the add-in task pane or anywhere in the host application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="8205e-119">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="8205e-119">Error Messages</span></span>

<span data-ttu-id="8205e-120">В таблице ниже перечислены ошибки, которые может возвращать API.</span><span class="sxs-lookup"><span data-stu-id="8205e-120">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="8205e-121">error.code</span><span class="sxs-lookup"><span data-stu-id="8205e-121">error.code</span></span> | <span data-ttu-id="8205e-122">error.message</span><span class="sxs-lookup"><span data-stu-id="8205e-122">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="8205e-123">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="8205e-123">InvalidArgument</span></span> |<span data-ttu-id="8205e-124">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="8205e-124">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="8205e-125">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="8205e-125">InvalidRequest</span></span>  |<span data-ttu-id="8205e-126">Не удается обработать запрос.</span><span class="sxs-lookup"><span data-stu-id="8205e-126">Cannot process the request.</span></span>|
|<span data-ttu-id="8205e-127">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="8205e-127">InvalidReference</span></span>|<span data-ttu-id="8205e-128">Эта ссылка недопустима для текущей операции.</span><span class="sxs-lookup"><span data-stu-id="8205e-128">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="8205e-129">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="8205e-129">InvalidBinding</span></span>  |<span data-ttu-id="8205e-130">Эта привязка объектов недопустима из-за предыдущих обновлений.</span><span class="sxs-lookup"><span data-stu-id="8205e-130">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="8205e-131">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="8205e-131">InvalidSelection</span></span>|<span data-ttu-id="8205e-132">Выбранный фрагмент недопустим для этой операции.</span><span class="sxs-lookup"><span data-stu-id="8205e-132">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="8205e-133">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="8205e-133">Unauthenticated</span></span> |<span data-ttu-id="8205e-134">Требуемые сведения о проверке подлинности отсутствуют или недопустимы.</span><span class="sxs-lookup"><span data-stu-id="8205e-134">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="8205e-135">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="8205e-135">AccessDenied</span></span> |<span data-ttu-id="8205e-136">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="8205e-136">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="8205e-137">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="8205e-137">ItemNotFound</span></span> |<span data-ttu-id="8205e-138">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="8205e-138">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="8205e-139">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="8205e-139">ActivityLimitReached</span></span>|<span data-ttu-id="8205e-140">Достигнут предел действий.</span><span class="sxs-lookup"><span data-stu-id="8205e-140">Activity limit has been reached.</span></span>|
|<span data-ttu-id="8205e-141">GeneralException</span><span class="sxs-lookup"><span data-stu-id="8205e-141">GeneralException</span></span>|<span data-ttu-id="8205e-142">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="8205e-142">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="8205e-143">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="8205e-143">NotImplemented</span></span>  |<span data-ttu-id="8205e-144">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="8205e-144">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="8205e-145">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="8205e-145">ServiceNotAvailable</span></span>|<span data-ttu-id="8205e-146">Служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="8205e-146">The service is unavailable.</span></span>|
|<span data-ttu-id="8205e-147">Conflict</span><span class="sxs-lookup"><span data-stu-id="8205e-147">Conflict</span></span>|<span data-ttu-id="8205e-148">Запрос не удалось обработать из-за конфликта.</span><span class="sxs-lookup"><span data-stu-id="8205e-148">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="8205e-149">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="8205e-149">ItemAlreadyExists</span></span>|<span data-ttu-id="8205e-150">Создаваемый ресурс уже существует.</span><span class="sxs-lookup"><span data-stu-id="8205e-150">The resource being created already exists.</span></span>|
|<span data-ttu-id="8205e-151">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="8205e-151">UnsupportedOperation</span></span>|<span data-ttu-id="8205e-152">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="8205e-152">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="8205e-153">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="8205e-153">RequestAborted</span></span>|<span data-ttu-id="8205e-154">Запрос прерван во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="8205e-154">The request was aborted during run time.</span></span>|
|<span data-ttu-id="8205e-155">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="8205e-155">ApiNotAvailable</span></span>|<span data-ttu-id="8205e-156">Запрашиваемый интерфейс API недоступен.</span><span class="sxs-lookup"><span data-stu-id="8205e-156">The requested API is not available.</span></span>|
|<span data-ttu-id="8205e-157">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="8205e-157">InsertDeleteConflict</span></span>|<span data-ttu-id="8205e-158">Операция вставки или удаления привела к конфликту.</span><span class="sxs-lookup"><span data-stu-id="8205e-158">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="8205e-159">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="8205e-159">InvalidOperation</span></span>|<span data-ttu-id="8205e-160">Выполняемая операция недопустима для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8205e-160">The operation attempted is invalid on the object.</span></span>|

## <a name="see-also"></a><span data-ttu-id="8205e-161">См. также</span><span class="sxs-lookup"><span data-stu-id="8205e-161">See also</span></span>

- [<span data-ttu-id="8205e-162">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="8205e-162">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8205e-163">Объект OfficeExtension.Error (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="8205e-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error)
