---
title: Обработка ошибок с помощью API JavaScript для Excel
description: Узнайте о логике обработки ошибок API JavaScript для Excel, чтобы учесть ошибки во время работы.
ms.date: 01/06/2021
localization_priority: Normal
ms.openlocfilehash: fd863e9783336ba9121312ba06aae03330d57562
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789123"
---
# <a name="error-handling-with-the-excel-javascript-api"></a><span data-ttu-id="4cdac-103">Обработка ошибок с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="4cdac-103">Error handling with the Excel JavaScript API</span></span>

<span data-ttu-id="4cdac-p101">При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.</span><span class="sxs-lookup"><span data-stu-id="4cdac-p101">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors. Doing so is critical, due to the asynchronous nature of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="4cdac-106">Дополнительные сведения о методе и асинхронном характере API JavaScript для Excel см. в объектной модели JavaScript для Excel в `sync()` [надстройки Office.](excel-add-ins-core-concepts.md)</span><span class="sxs-lookup"><span data-stu-id="4cdac-106">For more information about the `sync()` method and the asynchronous nature of Excel JavaScript API, see [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="4cdac-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="4cdac-107">Best practices</span></span>

<span data-ttu-id="4cdac-p102">В примерах кода в этой документации вы заметите, что каждый вызов `Excel.run` сопровождается оператором `catch`, что позволяет перехватывать все ошибки, возникающие в `Excel.run`. Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="4cdac-p102">Throughout the code samples in this documentation, you'll notice that every call to `Excel.run` is accompanied by a `catch` statement to catch any errors that occur within the `Excel.run`. We recommend that you use the same pattern when you build an add-in using the Excel JavaScript APIs.</span></span>

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

## <a name="api-errors"></a><span data-ttu-id="4cdac-110">Ошибки API</span><span class="sxs-lookup"><span data-stu-id="4cdac-110">API errors</span></span>

<span data-ttu-id="4cdac-111">Если не удается выполнить запрос API JavaScript для Excel, API возвращает объект error, содержащий следующие свойства:</span><span class="sxs-lookup"><span data-stu-id="4cdac-111">When an Excel JavaScript API request fails to run successfully, the API returns an error object that contains the following properties:</span></span>

- <span data-ttu-id="4cdac-p103">**code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.</span><span class="sxs-lookup"><span data-stu-id="4cdac-p103">**code**:  The `code` property of an error message contains a string that is part of the `OfficeExtension.ErrorCodes` or `Excel.ErrorCodes` list. For example, the error code "InvalidReference" indicates that the reference is not valid for the specified operation. Error codes are not localized.</span></span>

- <span data-ttu-id="4cdac-115">**message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке.</span><span class="sxs-lookup"><span data-stu-id="4cdac-115">**message**: The `message` property of an error message contains a summary of the error in the localized string.</span></span> <span data-ttu-id="4cdac-116">Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.</span><span class="sxs-lookup"><span data-stu-id="4cdac-116">The error message is not intended for consumption by end users; you should use the error code and appropriate business logic to determine the error message that your add-in shows to end users.</span></span>

- <span data-ttu-id="4cdac-117">**debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.</span><span class="sxs-lookup"><span data-stu-id="4cdac-117">**debugInfo**: When present, the `debugInfo` property of the error message provides additional information that you can use to understand the root cause of the error.</span></span>

> [!NOTE]
> <span data-ttu-id="4cdac-118">Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере.</span><span class="sxs-lookup"><span data-stu-id="4cdac-118">If you use `console.log()` to print error messages to the console, those messages will only be visible on the server.</span></span> <span data-ttu-id="4cdac-119">Конечные пользователи не будут видеть эти сообщения об ошибках в области задач надстройки или где-либо в приложении Office.</span><span class="sxs-lookup"><span data-stu-id="4cdac-119">End users will not see those error messages in the add-in task pane or anywhere in the Office application.</span></span>

## <a name="error-messages"></a><span data-ttu-id="4cdac-120">Сообщения об ошибках</span><span class="sxs-lookup"><span data-stu-id="4cdac-120">Error Messages</span></span>

<span data-ttu-id="4cdac-121">В таблице ниже перечислены ошибки, которые может возвращать API.</span><span class="sxs-lookup"><span data-stu-id="4cdac-121">The following table is a list of errors that the API may return.</span></span>

|<span data-ttu-id="4cdac-122">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4cdac-122">Error code</span></span> | <span data-ttu-id="4cdac-123">Сообщение об ошибке</span><span class="sxs-lookup"><span data-stu-id="4cdac-123">Error message</span></span> |
|:----------|:--------------|
|`AccessDenied` |<span data-ttu-id="4cdac-124">Вы не можете выполнить запрашиваемую операцию.</span><span class="sxs-lookup"><span data-stu-id="4cdac-124">You cannot perform the requested operation.</span></span>|
|`ActivityLimitReached`|<span data-ttu-id="4cdac-125">Достигнут предел действий.</span><span class="sxs-lookup"><span data-stu-id="4cdac-125">Activity limit has been reached.</span></span>|
|`ApiNotAvailable`|<span data-ttu-id="4cdac-126">Запрашиваемый интерфейс API недоступен.</span><span class="sxs-lookup"><span data-stu-id="4cdac-126">The requested API is not available.</span></span>|
|`ApiNotFound`|<span data-ttu-id="4cdac-127">Не удалось найти API, который вы пытаетесь использовать.</span><span class="sxs-lookup"><span data-stu-id="4cdac-127">The API you are trying to use could not be found.</span></span> <span data-ttu-id="4cdac-128">Она может быть доступна в более новой версии Excel.</span><span class="sxs-lookup"><span data-stu-id="4cdac-128">It may be available in a newer version of Excel.</span></span> <span data-ttu-id="4cdac-129">Дополнительные сведения см. в статье наборов требований [API JavaScript](../reference/requirement-sets/excel-api-requirement-sets.md) для Excel.</span><span class="sxs-lookup"><span data-stu-id="4cdac-129">See the [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) article for more information.</span></span>|
|`BadPassword`|<span data-ttu-id="4cdac-130">Пароль, который вы предоставили, неверен.</span><span class="sxs-lookup"><span data-stu-id="4cdac-130">The password you supplied is incorrect.</span></span>|
|`Conflict`|<span data-ttu-id="4cdac-131">Запрос не удалось обработать из-за конфликта.</span><span class="sxs-lookup"><span data-stu-id="4cdac-131">Request could not be processed because of a conflict.</span></span>|
|`ContentLengthRequired`|<span data-ttu-id="4cdac-132">Отсутствует `Content-length` заголок HTTP.</span><span class="sxs-lookup"><span data-stu-id="4cdac-132">A `Content-length` HTTP header is missing.</span></span>|
|`GeneralException`|<span data-ttu-id="4cdac-133">При обработке запроса возникла внутренняя ошибка.</span><span class="sxs-lookup"><span data-stu-id="4cdac-133">There was an internal error while processing the request.</span></span>|
|`InactiveWorkbook`|<span data-ttu-id="4cdac-134">Операция не удалась из-за того, что открыто несколько книг, а книга, вызванная этим API, теряет фокус.</span><span class="sxs-lookup"><span data-stu-id="4cdac-134">The operation failed because multiple workbooks are open and the workbook being called by this API has lost focus.</span></span>|
|`InsertDeleteConflict`|<span data-ttu-id="4cdac-135">Операция вставки или удаления привела к конфликту.</span><span class="sxs-lookup"><span data-stu-id="4cdac-135">The insert or delete operation attempted resulted in a conflict.</span></span>|
|`InvalidArgument` |<span data-ttu-id="4cdac-136">Аргумент недопустим, отсутствует или имеет неправильный формат.</span><span class="sxs-lookup"><span data-stu-id="4cdac-136">The argument is invalid or missing or has an incorrect format.</span></span>|
|`InvalidBinding`  |<span data-ttu-id="4cdac-137">Эта привязка объектов недопустима из-за предыдущих обновлений.</span><span class="sxs-lookup"><span data-stu-id="4cdac-137">This object binding is no longer valid due to previous updates.</span></span>|
|`InvalidOperation`|<span data-ttu-id="4cdac-138">Выполняемая операция недопустима для этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4cdac-138">The operation attempted is invalid on the object.</span></span>|
|`InvalidReference`|<span data-ttu-id="4cdac-139">Эта ссылка недопустима для текущей операции.</span><span class="sxs-lookup"><span data-stu-id="4cdac-139">This reference is not valid for the current operation.</span></span>|
|`InvalidRequest`  |<span data-ttu-id="4cdac-140">Не удается обработать запрос.</span><span class="sxs-lookup"><span data-stu-id="4cdac-140">Cannot process the request.</span></span>|
|`InvalidSelection`|<span data-ttu-id="4cdac-141">Выбранный фрагмент недопустим для этой операции.</span><span class="sxs-lookup"><span data-stu-id="4cdac-141">The current selection is invalid for this operation.</span></span>|
|`ItemAlreadyExists`|<span data-ttu-id="4cdac-142">Создаваемый ресурс уже существует.</span><span class="sxs-lookup"><span data-stu-id="4cdac-142">The resource being created already exists.</span></span>|
|`ItemNotFound` |<span data-ttu-id="4cdac-143">Запрашиваемый ресурс не существует.</span><span class="sxs-lookup"><span data-stu-id="4cdac-143">The requested resource doesn't exist.</span></span>|
|`NonBlankCellOffSheet`|<span data-ttu-id="4cdac-144">Запрос на вставку новых ячеек не может быть выполнен, так как он будет отставлять непустые ячейки с конца таблицы.</span><span class="sxs-lookup"><span data-stu-id="4cdac-144">The request to insert new cells can't be completed because it would push non-empty cells off the end of the worksheet.</span></span> <span data-ttu-id="4cdac-145">Эти непустые ячейки могут выглядеть пустыми, но иметь пустые значения, форматирование или формулу.</span><span class="sxs-lookup"><span data-stu-id="4cdac-145">These non-empty cells might appear empty but have blank values, some formatting, or a formula.</span></span> <span data-ttu-id="4cdac-146">Удалите достаточно строк или столбцов, чтобы уместить место для вставки, а затем попробуйте еще раз.</span><span class="sxs-lookup"><span data-stu-id="4cdac-146">Delete enough rows or columns to make room for what you want to insert and then try again.</span></span>|
|`NotImplemented`|<span data-ttu-id="4cdac-147">Запрашиваемая функция не реализована.</span><span class="sxs-lookup"><span data-stu-id="4cdac-147">The requested feature isn't implemented.</span></span>|
|`RangeExceedsLimit`|<span data-ttu-id="4cdac-148">Число ячеок в диапазоне превысило максимальное поддерживаемые числа.</span><span class="sxs-lookup"><span data-stu-id="4cdac-148">The cell count in the range has exceeded the maximum supported number.</span></span> <span data-ttu-id="4cdac-149">Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4cdac-149">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>|
|`RequestAborted`|<span data-ttu-id="4cdac-150">Запрос прерван во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="4cdac-150">The request was aborted during run time.</span></span>|
|`RequestPayloadSizeLimitExceeded`|<span data-ttu-id="4cdac-151">Размер полезной нагрузки запроса превысил ограничение.</span><span class="sxs-lookup"><span data-stu-id="4cdac-151">The request payload size has exceeded the limit.</span></span> <span data-ttu-id="4cdac-152">Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4cdac-152">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span> <br><br><span data-ttu-id="4cdac-153">Эта ошибка возникает только в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4cdac-153">This error only occurs in Excel on the web.</span></span>|
|`ResponsePayloadSizeLimitExceeded`|<span data-ttu-id="4cdac-154">Размер полезной нагрузки отклика превысил ограничение.</span><span class="sxs-lookup"><span data-stu-id="4cdac-154">The response payload size has exceeded the limit.</span></span> <span data-ttu-id="4cdac-155">Дополнительные [сведения см.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) в статье об ограничениях ресурсов и оптимизации производительности надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="4cdac-155">See the [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins) article for more information.</span></span>  <br><br><span data-ttu-id="4cdac-156">Эта ошибка возникает только в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="4cdac-156">This error only occurs in Excel on the web.</span></span>|
|`ServiceNotAvailable`|<span data-ttu-id="4cdac-157">Служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="4cdac-157">The service is unavailable.</span></span>|
|`Unauthenticated` |<span data-ttu-id="4cdac-158">Требуемые сведения о проверке подлинности отсутствуют или недопустимы.</span><span class="sxs-lookup"><span data-stu-id="4cdac-158">Required authentication information is either missing or invalid.</span></span>|
|`UnsupportedOperation`|<span data-ttu-id="4cdac-159">Выполняемая операция не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4cdac-159">The operation being attempted is not supported.</span></span>|
|`UnsupportedSheet`|<span data-ttu-id="4cdac-160">Этот тип листа не поддерживает эту операцию, так как он является листом макроса или диаграммы.</span><span class="sxs-lookup"><span data-stu-id="4cdac-160">This sheet type does not support this operation, since it is a Macro or Chart sheet.</span></span>|

## <a name="see-also"></a><span data-ttu-id="4cdac-161">См. также</span><span class="sxs-lookup"><span data-stu-id="4cdac-161">See also</span></span>

- [<span data-ttu-id="4cdac-162">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="4cdac-162">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4cdac-163">Объект OfficeExtension.Error (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="4cdac-163">OfficeExtension.Error object (JavaScript API for Excel)</span></span>](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
