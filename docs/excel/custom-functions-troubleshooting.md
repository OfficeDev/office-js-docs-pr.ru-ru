---
ms.date: 03/30/2021
description: Устранение распространенных проблем с пользовательскими функциями Excel.
title: Устранение проблем в пользовательских функциях
localization_priority: Normal
ms.openlocfilehash: e79b2f8ee8abccda2b34821761bab65592a90218
ms.sourcegitcommit: 074526a6dca8381dbdabf2705474c5ae6753b829
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51506142"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="7ea06-103">Устранение проблем в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="7ea06-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="7ea06-104">При разработке пользовательских функций могут возникать ошибки в продукте при создании и тестировании функций.</span><span class="sxs-lookup"><span data-stu-id="7ea06-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="7ea06-105">Для устранения проблем можно [включить ведение журнала в среде выполнения для регистрации ошибок](#enable-runtime-logging) и ознакомиться с [собственными сообщениями об ошибках Excel](#check-for-excel-error-messages).</span><span class="sxs-lookup"><span data-stu-id="7ea06-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="7ea06-106">Проверьте также на наличие распространенных ошибок, например [оставление неразрешенных обещаний](#ensure-promises-return).</span><span class="sxs-lookup"><span data-stu-id="7ea06-106">Also, check for common mistakes such as [leaving promises unresolved](#ensure-promises-return).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="7ea06-107">Включение ведения журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="7ea06-107">Enable runtime logging</span></span>

<span data-ttu-id="7ea06-108">Если вы тестируете надстройку в Office для Windows, следует [включить ведение журнала среды выполнения](../testing/runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="7ea06-108">If you're testing your add-in in Office on Windows, you should [enable runtime logging](../testing/runtime-logging.md).</span></span> <span data-ttu-id="7ea06-109">Ведение журнала в среде выполнения отправляет операторы `console.log` в отдельный файл журнала, созданный для выявления проблем.</span><span class="sxs-lookup"><span data-stu-id="7ea06-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="7ea06-110">Операторы охватывают разнообразные ошибки, включая относящиеся к XML-файлу манифеста надстройки, условиям среды выполнения или установке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="7ea06-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span> <span data-ttu-id="7ea06-111">Дополнительные сведения о ведении журнала среды выполнения см. в статье [Отладка надстройки с помощью журнала среды выполнения](../testing/runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="7ea06-111">For more information about runtime logging, see [Debug your add-in with runtime logging](../testing/runtime-logging.md).</span></span>

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="7ea06-112">Проверка наличия сообщений об ошибках Excel</span><span class="sxs-lookup"><span data-stu-id="7ea06-112">Check for Excel error messages</span></span>

<span data-ttu-id="7ea06-113">В Excel есть несколько встроенных сообщений об ошибках, возвращаемых в ячейку при возникновении ошибки вычислений.</span><span class="sxs-lookup"><span data-stu-id="7ea06-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="7ea06-114">Для пользовательских функций используются только следующие сообщения об ошибках: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` и `#BUSY!`.</span><span class="sxs-lookup"><span data-stu-id="7ea06-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="7ea06-115">В общем случае такие ошибки уже будут вам знакомы после работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="7ea06-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="7ea06-116">Есть только несколько исключений, специфических для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="7ea06-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="7ea06-117">Ошибка `#NAME` обычно означает проблему с регистрацией функции.</span><span class="sxs-lookup"><span data-stu-id="7ea06-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="7ea06-118">Ошибка `#N/A` может также указывать на то, что зарегистрированную функцию не удалось выполнить.</span><span class="sxs-lookup"><span data-stu-id="7ea06-118">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="7ea06-119">Как правило, так происходит из-за отсутствия команды `CustomFunctions.associate`.</span><span class="sxs-lookup"><span data-stu-id="7ea06-119">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="7ea06-120">Ошибка `#VALUE` обычно связана с файлом сценария функций.</span><span class="sxs-lookup"><span data-stu-id="7ea06-120">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="7ea06-121">Ошибка `#REF!` может указывать на то, что имя функции совпадает с именем другой функции, которая уже есть в надстройке.</span><span class="sxs-lookup"><span data-stu-id="7ea06-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="7ea06-122">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="7ea06-122">Clear the Office cache</span></span>

<span data-ttu-id="7ea06-123">Office помещает сведения о пользовательских функциях в кэш.</span><span class="sxs-lookup"><span data-stu-id="7ea06-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="7ea06-124">Иногда при разработке и многократной повторной загрузке надстройки с пользовательскими функциями изменения могут не отображаться.</span><span class="sxs-lookup"><span data-stu-id="7ea06-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="7ea06-125">Это можно исправить, очистив кэш Office.</span><span class="sxs-lookup"><span data-stu-id="7ea06-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="7ea06-126">Дополнительные сведения см. в статье [Очистка кэша Office](../testing/clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="7ea06-126">For more information, see [Clear the Office cache](../testing/clear-cache.md).</span></span>

## <a name="common-problems-and-solutions"></a><span data-ttu-id="7ea06-127">Общие проблемы и решения</span><span class="sxs-lookup"><span data-stu-id="7ea06-127">Common problems and solutions</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exemption"></a><span data-ttu-id="7ea06-128">Не может открыть надстройку из localhost: используйте локальное освобождение от циклов.</span><span class="sxs-lookup"><span data-stu-id="7ea06-128">Can't open add-in from localhost: Use a local loopback exemption</span></span>

<span data-ttu-id="7ea06-129">Если вы видите ошибку "Мы не можем открыть эту надстройку из localhost", необходимо включить локальное освобождение от циклов.</span><span class="sxs-lookup"><span data-stu-id="7ea06-129">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exemption.</span></span> <span data-ttu-id="7ea06-130">Подробные сведения о том, как это сделать, см. в [этой статье службы поддержки Майкрософт](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span><span class="sxs-lookup"><span data-stu-id="7ea06-130">For details on how to do this, see [this Microsoft support article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span></span>

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a><span data-ttu-id="7ea06-131">В журнале среды выполнения возникает сообщение об ошибке "TypeError: сетевой запрос не выполнен" в Excel для Windows</span><span class="sxs-lookup"><span data-stu-id="7ea06-131">Runtime logging reports "TypeError: Network request failed" on Excel on Windows</span></span>

<span data-ttu-id="7ea06-132">Если в вашем [журнале среды выполнения](custom-functions-troubleshooting.md#enable-runtime-logging) отображается ошибка "TypeError: сетевой запрос не выполнен" при вызове сервера localhost, требуется включить исключение локального замыкания на себя.</span><span class="sxs-lookup"><span data-stu-id="7ea06-132">If you see the error "TypeError: Network request failed" in your [runtime log](custom-functions-troubleshooting.md#enable-runtime-logging) while making calls to your localhost server, you'll need to enable a local loopback exception.</span></span> <span data-ttu-id="7ea06-133">Дополнительные сведения о том, как это сделать, см. в разделе *Вариант 2* [этой статьи от службы поддержки Майкрософт](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).</span><span class="sxs-lookup"><span data-stu-id="7ea06-133">For details on how to do this, see *Option #2* in [this Microsoft support article](https://support.microsoft.com/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="7ea06-134">Проверка возвращения обещаний</span><span class="sxs-lookup"><span data-stu-id="7ea06-134">Ensure promises return</span></span>

<span data-ttu-id="7ea06-135">Если Excel ожидает завершения выполнения пользовательской функции, выводится сообщение #BUSY!</span><span class="sxs-lookup"><span data-stu-id="7ea06-135">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="7ea06-136">в ячейке.</span><span class="sxs-lookup"><span data-stu-id="7ea06-136">in the cell.</span></span> <span data-ttu-id="7ea06-137">Если код пользовательской функции возвращает обещание, но это обещание не возвращает результат, Excel продолжит отображать сообщение `#BUSY!`.</span><span class="sxs-lookup"><span data-stu-id="7ea06-137">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing `#BUSY!`.</span></span> <span data-ttu-id="7ea06-138">Проверьте свои функции, чтобы убедиться, что все обещания правильно возвращают результат в ячейку.</span><span class="sxs-lookup"><span data-stu-id="7ea06-138">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="7ea06-139">Ошибка: сервер разработки уже работает на порту 3000</span><span class="sxs-lookup"><span data-stu-id="7ea06-139">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="7ea06-140">Иногда во время работы `npm start` отображается ошибка с сообщением о том, что сервер разработки уже работает на порту 3000 (или на любом другом порту, который используется надстройкой).</span><span class="sxs-lookup"><span data-stu-id="7ea06-140">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="7ea06-141">Можно остановить сервер разработки, выполнив `npm stop` или закрыв окно Node.js.</span><span class="sxs-lookup"><span data-stu-id="7ea06-141">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="7ea06-142">В некоторых случаях для остановки работы сервера разработчиков может занять несколько минут.</span><span class="sxs-lookup"><span data-stu-id="7ea06-142">In some cases, it can take a few minutes for the dev server to stop running.</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="7ea06-143">Мои функции не загружаются: свяжите функции</span><span class="sxs-lookup"><span data-stu-id="7ea06-143">My functions won't load: associate functions</span></span>

<span data-ttu-id="7ea06-144">Если вы не зарегистрировали JSON и создали собственные метаданные JSON, может появиться сообщение об ошибке `#VALUE!` или уведомление о том, что надстройку не удается загрузить.</span><span class="sxs-lookup"><span data-stu-id="7ea06-144">In cases where your JSON has not been registered and you have authored your own JSON metadata, you may see a `#VALUE!` error or receive a notification that your add-in cannot be loaded.</span></span> <span data-ttu-id="7ea06-145">Это обычно означает, что необходимо связать каждую пользовательскую функцию с ее свойством `id`, указанным в [файле метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="7ea06-145">This usually means you need to associate each custom function with its `id` property specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="7ea06-146">Это выполняется с помощью метода `CustomFunctions.associate()`.</span><span class="sxs-lookup"><span data-stu-id="7ea06-146">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="7ea06-147">Обычно вызов этого метода выполняется после каждой функции или в конце файла скрипта.</span><span class="sxs-lookup"><span data-stu-id="7ea06-147">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="7ea06-148">Если пользовательская функция не связана, она не будет работать.</span><span class="sxs-lookup"><span data-stu-id="7ea06-148">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="7ea06-149">В примере ниже показана функция добавления, за которой следует имя функции `add`, связанное с соответствующим идентификатором JSON `ADD`.</span><span class="sxs-lookup"><span data-stu-id="7ea06-149">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="7ea06-150">Дополнительные сведения об этом процессе см. в ссылке [Associating function names with JSON metadata.](../excel/custom-functions-json.md#associating-function-names-with-json-metadata)</span><span class="sxs-lookup"><span data-stu-id="7ea06-150">For more information on this process, see [Associating function names with JSON metadata](../excel/custom-functions-json.md#associating-function-names-with-json-metadata).</span></span>

## <a name="known-issues"></a><span data-ttu-id="7ea06-151">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="7ea06-151">Known issues</span></span>

<span data-ttu-id="7ea06-152">Известные проблемы отслеживаются и сообщаются в репозитории [Excel Custom Functions GitHub.](https://github.com/OfficeDev/Excel-Custom-Functions/issues)</span><span class="sxs-lookup"><span data-stu-id="7ea06-152">Known issues are tracked and reported in the [Excel Custom Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="7ea06-153">Обратная связь</span><span class="sxs-lookup"><span data-stu-id="7ea06-153">Reporting feedback</span></span>

<span data-ttu-id="7ea06-154">Если у вас возникают проблемы, не описанные в этой статье, сообщите нам.</span><span class="sxs-lookup"><span data-stu-id="7ea06-154">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="7ea06-155">Сообщить о проблемах можно двумя способами.</span><span class="sxs-lookup"><span data-stu-id="7ea06-155">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="7ea06-156">В Excel для Windows или Mac</span><span class="sxs-lookup"><span data-stu-id="7ea06-156">In Excel on Windows or Mac</span></span>

<span data-ttu-id="7ea06-157">При использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel.</span><span class="sxs-lookup"><span data-stu-id="7ea06-157">If using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="7ea06-158">Для этого выберите **Файл -> Отзыв -> Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="7ea06-158">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="7ea06-159">Отправка нахмуренного смайлика предоставит необходимые журналы для понимания проблемы, на которую вы указываете.</span><span class="sxs-lookup"><span data-stu-id="7ea06-159">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="7ea06-160">В GitHub</span><span class="sxs-lookup"><span data-stu-id="7ea06-160">In Github</span></span>

<span data-ttu-id="7ea06-161">Вы можете сообщить о возникшей проблеме либо с помощью функции отзыва о содержимом внизу любой страницы с документацией или [сообщить о новой проблеме непосредственно в репозитории пользовательских функций](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="7ea06-161">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="7ea06-162">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7ea06-162">Next steps</span></span>
<span data-ttu-id="7ea06-163">Узнайте, как [создавать пользовательские функции, совместимые с функциями XLL, определенными пользователями](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="7ea06-163">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7ea06-164">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="7ea06-164">See also</span></span>

* [<span data-ttu-id="7ea06-165">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7ea06-165">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="7ea06-166">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="7ea06-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
