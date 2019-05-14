---
ms.date: 05/08/2019
description: Устранение распространенных проблем в пользовательских функциях Excel.
title: Устранение проблем в пользовательских функциях
localization_priority: Priority
ms.openlocfilehash: 999b1fb9b89050ab5c6bcf87e1aac9d2fce13702
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952056"
---
# <a name="troubleshoot-custom-functions"></a><span data-ttu-id="76f90-103">Устранение проблем в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="76f90-103">Troubleshoot custom functions</span></span>

<span data-ttu-id="76f90-104">При разработке пользовательских функций могут возникать ошибки в продукте при создании и тестировании функций.</span><span class="sxs-lookup"><span data-stu-id="76f90-104">When developing custom functions, you may encounter errors in the product while creating and testing your functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="76f90-105">Для устранения проблем можно [включить ведение журнала в среде выполнения для регистрации ошибок](#enable-runtime-logging) и ознакомиться с [собственными сообщениями об ошибках Excel](#check-for-excel-error-messages).</span><span class="sxs-lookup"><span data-stu-id="76f90-105">To resolve issues, you can [enable runtime logging to capture errors](#enable-runtime-logging) and refer to [Excel's native error messages](#check-for-excel-error-messages).</span></span> <span data-ttu-id="76f90-106">Проверьте также отсутствие распространенных ошибок, например [оставление неразрешенных обещаний](#ensure-promises-return) и невыполнение [связывания функций](#my-functions-wont-load-associate-functions).</span><span class="sxs-lookup"><span data-stu-id="76f90-106">Also, check for common mistakes such as [leaving promises unresolved](#ensure-promises-return) and forgetting to [associate your functions](#my-functions-wont-load-associate-functions).</span></span>

## <a name="enable-runtime-logging"></a><span data-ttu-id="76f90-107">Включение ведения журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="76f90-107">Enable runtime logging</span></span>

<span data-ttu-id="76f90-108">Если вы тестируете надстройку в Office для Windows, следует [включить ведение журнала в среде выполнения](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="76f90-108">If you are testing your add-in in Office on Windows, you should [enable runtime logging](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span> <span data-ttu-id="76f90-109">Ведение журнала в среде выполнения отправляет операторы `console.log` в отдельный файл журнала, созданный для выявления проблем.</span><span class="sxs-lookup"><span data-stu-id="76f90-109">Runtime logging delivers `console.log` statements to a separate log file you create to help you uncover issues.</span></span> <span data-ttu-id="76f90-110">Операторы охватывают разнообразные ошибки, включая относящиеся к XML-файлу манифеста надстройки, условиям среды выполнения или установке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="76f90-110">The statements cover a variety of errors, including errors pertaining to your add-in's XML manifest file, runtime conditions, or installation of your custom functions.</span></span>  <span data-ttu-id="76f90-111">Дополнительные сведения о ведении журнала в среде выполнения см. в статье [Отладка надстройки с помощью журнала среды выполнения](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="76f90-111">For more information about runtime logging, see [Use runtime logging to debug your add-in](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in).</span></span>  

### <a name="check-for-excel-error-messages"></a><span data-ttu-id="76f90-112">Проверка наличия сообщений об ошибках Excel</span><span class="sxs-lookup"><span data-stu-id="76f90-112">Check for Excel error messages</span></span>

<span data-ttu-id="76f90-113">В Excel есть несколько встроенных сообщений об ошибках, возвращаемых в ячейку при возникновении ошибки вычислений.</span><span class="sxs-lookup"><span data-stu-id="76f90-113">Excel has a number of built-in error messages which are returned to a cell if there is calculation error.</span></span> <span data-ttu-id="76f90-114">Для пользовательских функций используются только следующие сообщения об ошибках: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` и `#BUSY!`.</span><span class="sxs-lookup"><span data-stu-id="76f90-114">Custom functions only use the following error messages: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, and `#BUSY!`.</span></span>

<span data-ttu-id="76f90-115">В общем случае такие ошибки уже будут вам знакомы после работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="76f90-115">Generally, these errors correspond to the errors you might already be familiar with in Excel.</span></span> <span data-ttu-id="76f90-116">Есть только несколько исключений, специфических для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="76f90-116">The are only a few exceptions specific to custom functions, listed here:</span></span>

- <span data-ttu-id="76f90-117">Ошибка `#NAME` обычно означает проблему с регистрацией функции.</span><span class="sxs-lookup"><span data-stu-id="76f90-117">A `#NAME` error generally means there has been an issue registering your functions.</span></span>
- <span data-ttu-id="76f90-118">Ошибка `#VALUE` обычно связана с файлом сценария функций.</span><span class="sxs-lookup"><span data-stu-id="76f90-118">A `#VALUE` error typically indicates an error in the functions' script file.</span></span>
- <span data-ttu-id="76f90-119">Ошибка `#N/A` может также указывать на то, что зарегистрированную функцию не удалось выполнить.</span><span class="sxs-lookup"><span data-stu-id="76f90-119">A `#N/A` error is also maybe a sign that that function while registered could not be run.</span></span> <span data-ttu-id="76f90-120">Как правило, так происходит из-за отсутствия команды `CustomFunctions.associate`.</span><span class="sxs-lookup"><span data-stu-id="76f90-120">This is typically due to a missing `CustomFunctions.associate` command.</span></span>
- <span data-ttu-id="76f90-121">Ошибка `#REF!` может указывать на то, что имя функции совпадает с именем другой функции, которая уже есть в надстройке.</span><span class="sxs-lookup"><span data-stu-id="76f90-121">A `#REF!` error may indicate that your function name is the same as a function name in an add-in that already exists.</span></span>

## <a name="clear-the-office-cache"></a><span data-ttu-id="76f90-122">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="76f90-122">Clear the Office cache</span></span>

<span data-ttu-id="76f90-123">Office помещает сведения о пользовательских функциях в кэш.</span><span class="sxs-lookup"><span data-stu-id="76f90-123">Information about custom functions is cached by Office.</span></span> <span data-ttu-id="76f90-124">Иногда при разработке и многократной повторной загрузке надстройки с пользовательскими функциями изменения могут не отображаться.</span><span class="sxs-lookup"><span data-stu-id="76f90-124">Sometimes while developing and repeatedly reloading an add-in with custom functions your changes may not appear.</span></span> <span data-ttu-id="76f90-125">Это можно исправить, очистив кэш Office.</span><span class="sxs-lookup"><span data-stu-id="76f90-125">You can fix this by clearing the Office cache.</span></span> <span data-ttu-id="76f90-126">Дополнительные сведения см. в разделе «Очистить кэш Office» в статье [Проблемы с проверкой и устранением неполадок манифеста](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span><span class="sxs-lookup"><span data-stu-id="76f90-126">For more information, see the "Clear the Office cache" section in the article [Validate and troubleshoot issues with your manifest](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)</span></span>

## <a name="common-issues"></a><span data-ttu-id="76f90-127">Распространенные проблемы</span><span class="sxs-lookup"><span data-stu-id="76f90-127">Common issues</span></span>

### <a name="my-functions-wont-load-associate-functions"></a><span data-ttu-id="76f90-128">Мои функции не загружаются: свяжите функции</span><span class="sxs-lookup"><span data-stu-id="76f90-128">My functions won't load: associate functions</span></span>

<span data-ttu-id="76f90-129">В файле скрипта пользовательских функций необходимо связать каждую пользовательскую функцию с ее идентификатором, указанным в [файле метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="76f90-129">In your custom functions' script file, you need to associate each custom function with its ID specified in the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="76f90-130">Это выполняется с помощью метода `CustomFunctions.associate()`.</span><span class="sxs-lookup"><span data-stu-id="76f90-130">This is done by using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="76f90-131">Обычно вызов этого метода выполняется после каждой функции или в конце файла скрипта.</span><span class="sxs-lookup"><span data-stu-id="76f90-131">Typically this method call is made after each function or at the end of the script file.</span></span> <span data-ttu-id="76f90-132">Если пользовательская функция не связана, она не будет работать.</span><span class="sxs-lookup"><span data-stu-id="76f90-132">If a custom function is not associated, it will not work.</span></span>

<span data-ttu-id="76f90-133">В примере ниже показана функция добавления, за которой следует имя функции `add`, связанное с соответствующим идентификатором JSON `ADD`.</span><span class="sxs-lookup"><span data-stu-id="76f90-133">The following example shows an add function, followed by the function's name `add` being associated with the corresponding JSON id `ADD`.</span></span>

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

<span data-ttu-id="76f90-134">Дополнительные сведения об этом процессе см. в статье [Сопоставление имен функций с метаданными JSON](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="76f90-134">For more information on this process, see [Associating function names with json metadata](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata).</span></span>

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a><span data-ttu-id="76f90-135">Не удается открыть надстройку из узла localhost: используйте исключение локального замыкания на себя</span><span class="sxs-lookup"><span data-stu-id="76f90-135">Can't open add-in from localhost: use a local loopback exception</span></span>

<span data-ttu-id="76f90-136">Если отображается ошибка "Не удается открыть эту надстройку из localhost", необходимо включить исключение локального замыкания на себя.</span><span class="sxs-lookup"><span data-stu-id="76f90-136">If you see the error "We can't open this add-in from localhost," you will need to enable a local loopback exception.</span></span> <span data-ttu-id="76f90-137">Подробные сведения о том, как это сделать, см. в [этой статье службы поддержки Майкрософт](https://support.microsoft.com/ru-RU/help/4490419/local-loopback-exemption-does-not-work).</span><span class="sxs-lookup"><span data-stu-id="76f90-137">For details on how to do this, see [this Microsoft support article](https://support.microsoft.com/ru-RU/help/4490419/local-loopback-exemption-does-not-work).</span></span>

### <a name="ensure-promises-return"></a><span data-ttu-id="76f90-138">Проверка возвращения обещаний</span><span class="sxs-lookup"><span data-stu-id="76f90-138">Ensure promises return</span></span>

<span data-ttu-id="76f90-139">Если Excel ожидает завершения выполнения пользовательской функции, выводится сообщение #BUSY!</span><span class="sxs-lookup"><span data-stu-id="76f90-139">When Excel is waiting for a custom function to complete, it displays #BUSY!</span></span> <span data-ttu-id="76f90-140">в ячейке.</span><span class="sxs-lookup"><span data-stu-id="76f90-140">in the cell.</span></span> <span data-ttu-id="76f90-141">Если код пользовательской функции возвращает обещание, но это обещание не возвращает результат, Excel продолжит отображать сообщение #BUSY!.</span><span class="sxs-lookup"><span data-stu-id="76f90-141">If your custom function code returns a promise, but the promise does not return a result, Excel will continue showing #BUSY!.</span></span> <span data-ttu-id="76f90-142">Проверьте свои функции, чтобы убедиться, что все обещания правильно возвращают результат в ячейку.</span><span class="sxs-lookup"><span data-stu-id="76f90-142">Check your functions to make sure that any promises are properly returning a result to a cell.</span></span>

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a><span data-ttu-id="76f90-143">Ошибка: сервер разработки уже работает на порту 3000</span><span class="sxs-lookup"><span data-stu-id="76f90-143">Error: The dev server is already running on port 3000</span></span>

<span data-ttu-id="76f90-144">Иногда во время работы `npm start` отображается ошибка с сообщением о том, что сервер разработки уже работает на порту 3000 (или на любом другом порту, который используется надстройкой).</span><span class="sxs-lookup"><span data-stu-id="76f90-144">Sometimes when running `npm start` you may see an error that the dev server is already running on port 3000 (or whichever port your add-in uses).</span></span> <span data-ttu-id="76f90-145">Можно остановить сервер разработки, выполнив `npm stop` или закрыв окно Node.js.</span><span class="sxs-lookup"><span data-stu-id="76f90-145">You can stop the dev server by running `npm stop` or by closing the Node.js window.</span></span> <span data-ttu-id="76f90-146">Но в некоторых случаях проходит несколько минут, прежде чем сервер разработки действительно остановится.</span><span class="sxs-lookup"><span data-stu-id="76f90-146">But in some cases in can take a few minutes for the dev server to actually stop running.</span></span>

## <a name="reporting-feedback"></a><span data-ttu-id="76f90-147">Обратная связь</span><span class="sxs-lookup"><span data-stu-id="76f90-147">Reporting feedback</span></span>

<span data-ttu-id="76f90-148">Если у вас возникают проблемы, не описанные в этой статье, сообщите нам.</span><span class="sxs-lookup"><span data-stu-id="76f90-148">If you are encountering issues that aren't documented here, let us know.</span></span> <span data-ttu-id="76f90-149">Сообщить о проблемах можно двумя способами.</span><span class="sxs-lookup"><span data-stu-id="76f90-149">There are two ways to report issues.</span></span>

### <a name="in-excel-on-windows-or-mac"></a><span data-ttu-id="76f90-150">В Excel для Windows или Mac</span><span class="sxs-lookup"><span data-stu-id="76f90-150">In Excel on Windows or Mac</span></span>

<span data-ttu-id="76f90-151">При использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel.</span><span class="sxs-lookup"><span data-stu-id="76f90-151">If using Excel for Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="76f90-152">Для этого выберите **Файл -> Отзыв -> Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="76f90-152">To do this, select **File -> Feedback -> Send a Frown**.</span></span> <span data-ttu-id="76f90-153">Отправка нахмуренного смайлика предоставит необходимые журналы для понимания проблемы, на которую вы указываете.</span><span class="sxs-lookup"><span data-stu-id="76f90-153">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

### <a name="in-github"></a><span data-ttu-id="76f90-154">В GitHub</span><span class="sxs-lookup"><span data-stu-id="76f90-154">In Github</span></span>

<span data-ttu-id="76f90-155">Вы можете сообщить о возникшей проблеме либо с помощью функции отзыва о содержимом внизу любой страницы с документацией или [сообщить о новой проблеме непосредственно в репозитории пользовательских функций](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="76f90-155">Feel free to submit an issue you encounter either through the "Content feedback" feature at the bottom of any documentation page, or by [filing a new issue directly to the custom functions repository](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="76f90-156">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="76f90-156">Next steps</span></span>
<span data-ttu-id="76f90-157">Узнайте, как [отладить пользовательские функции](custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="76f90-157">Learn how to [debug your custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="76f90-158">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="76f90-158">See also</span></span>

* [<span data-ttu-id="76f90-159">Автогенерация метаданных для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="76f90-159">Custom functions metadata autogeneration</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="76f90-160">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="76f90-160">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="76f90-161">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="76f90-161">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="76f90-162">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="76f90-162">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="76f90-163">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="76f90-163">Create custom functions in Excel</span></span>](custom-functions-overview.md)
