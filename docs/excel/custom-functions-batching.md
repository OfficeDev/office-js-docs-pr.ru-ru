---
ms.date: 04/22/2019
description: Объедините пользовательские функции в пакет, чтобы сократить количество обращений к удаленной службе через сеть.
title: Пакетирование обращений пользовательских функций к удаленной службе
localization_priority: Priority
ms.openlocfilehash: 2e31d6aa212e27967448f07fdcb2bd024a7511f9
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356914"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a><span data-ttu-id="a5d58-103">Пакетирование обращений пользовательских функций к удаленной службе</span><span class="sxs-lookup"><span data-stu-id="a5d58-103">Batching custom function calls for a remote service</span></span>

<span data-ttu-id="a5d58-104">Если пользовательские функции обращаются к удаленной службе, можно использовать шаблон пакетирования для сокращения количества сетевых вызовов удаленной службы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-104">If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service.</span></span> <span data-ttu-id="a5d58-105">Для уменьшения объема сетевых операций можно объединить все вызовы в один вызов веб-службы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-105">To reduce network round trips you batch all the calls into a single call to the web service.</span></span> <span data-ttu-id="a5d58-106">Это идеальное решение при пересчете электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-106">This is ideal when the spreadsheet is recalculated.</span></span> <span data-ttu-id="a5d58-107">Например если пользователь обращается к вашей пользовательской функции в 100 ячейках электронной таблицы, а затем пересчитывает электронную таблицу, эта функция будет выполняться 100 раз и делать 100 сетевых вызовов.</span><span class="sxs-lookup"><span data-stu-id="a5d58-107">For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls.</span></span> <span data-ttu-id="a5d58-108">С помощью шаблона пакетирования эти вызовы можно объединить так, чтобы делать 100 расчетов в течение одного сетевого вызова.</span><span class="sxs-lookup"><span data-stu-id="a5d58-108">By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.</span></span>

## <a name="view-the-completed-sample"></a><span data-ttu-id="a5d58-109">Посмотреть готовый пример</span><span class="sxs-lookup"><span data-stu-id="a5d58-109">View the completed sample</span></span>

<span data-ttu-id="a5d58-110">Вы можете изучить эту статью и вставить примеры кода в свой проект.</span><span class="sxs-lookup"><span data-stu-id="a5d58-110">You can follow this article and paste the code examples into your own project.</span></span> <span data-ttu-id="a5d58-111">Например можно создать в Office проект пользовательской функции для TypeScript, а затем вставить в него весь код из этой статьи,</span><span class="sxs-lookup"><span data-stu-id="a5d58-111">For example, you can use yo office to create a new custom function project for TypeScript, then add all the code from this article to the project.</span></span> <span data-ttu-id="a5d58-112">а затем запустить код и посмотреть на результаты его работы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-112">you can then run the code and try it out.</span></span>

<span data-ttu-id="a5d58-113">Также можно загрузить или просмотреть готовый образец проекта на странице [Custom function batching pattern (Пакетирование пользовательских функций)](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span><span class="sxs-lookup"><span data-stu-id="a5d58-113">Also you can download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span></span> <span data-ttu-id="a5d58-114">Если вы хотите просмотреть код в целом, прежде чем читать дальше, посмотрите на [файл сценария](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).</span><span class="sxs-lookup"><span data-stu-id="a5d58-114">If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).</span></span>

## <a name="create-the-batching-pattern-in-this-article"></a><span data-ttu-id="a5d58-115">Создание шаблона пакетирования в этой статье</span><span class="sxs-lookup"><span data-stu-id="a5d58-115">Create the batching pattern in this article</span></span>

<span data-ttu-id="a5d58-116">Для реализации пакетирования пользовательских функций необходимо создать три основных раздела кода.</span><span class="sxs-lookup"><span data-stu-id="a5d58-116">To set up batching for your custom functions you'll need to write three main sections of code.</span></span>

1. <span data-ttu-id="a5d58-117">Push-операция для включения новой операции в пакет вызовов каждый раз, когда Excel вызывает пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="a5d58-117">A push operation to add a new operation to the batch of calls each time Excel calls your custom function.</span></span>
2. <span data-ttu-id="a5d58-118">Функция, которая делает удаленный запрос, когда пакет готов.</span><span class="sxs-lookup"><span data-stu-id="a5d58-118">A function to make the remote request when the batch is ready.</span></span>
3. <span data-ttu-id="a5d58-119">Код сервера для отклика на пакетный запрос, вычисления результатов всех операций и возвращения значений.</span><span class="sxs-lookup"><span data-stu-id="a5d58-119">Server code to respond to the batch request, calculate all of the operation results, and return the values.</span></span>

<span data-ttu-id="a5d58-120">В следующих разделах будет показано создание кода по одному примеру за раз.</span><span class="sxs-lookup"><span data-stu-id="a5d58-120">In the following sections you will be shown how to construct the code one example at a time.</span></span> <span data-ttu-id="a5d58-121">Добавьте каждый пример кода в файл functions.ts.</span><span class="sxs-lookup"><span data-stu-id="a5d58-121">You'll add each code example to your functions.ts file.</span></span> <span data-ttu-id="a5d58-122">Рекомендуем создавать пользовательские функции заново в вашей копии Office.</span><span class="sxs-lookup"><span data-stu-id="a5d58-122">It's recommended you create a brand new custom functions project using yo office.</span></span> <span data-ttu-id="a5d58-123">Для создания проекта обратитесь к статье [Начало разработки пользовательских функций Excel](../quickstarts/excel-custom-functions-quickstart.md) и используйте TypeScript вместо JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a5d58-123">To create a new project see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md) and use TypeScript instead of JavaScript.</span></span>

## <a name="batch-each-call-to-your-custom-function"></a><span data-ttu-id="a5d58-124">Включение в пакет каждого вызова пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="a5d58-124">Batch each call to your custom function</span></span>

<span data-ttu-id="a5d58-125">Ваши пользовательские функции вызывают удаленную службу для выполнения различных операций и вычисления требуемого результата.</span><span class="sxs-lookup"><span data-stu-id="a5d58-125">Your custom functions work by calling a remote service to perform the operation and calculate the result they need.</span></span> <span data-ttu-id="a5d58-126">Это дает возможность сохранения каждой запрашиваемой операции в пакете.</span><span class="sxs-lookup"><span data-stu-id="a5d58-126">This provides a way for them to store each requested operation into a batch.</span></span> <span data-ttu-id="a5d58-127">Далее вы узнаете, как создать функцию `_pushOperation` для пакетной обработки операций.</span><span class="sxs-lookup"><span data-stu-id="a5d58-127">Later you'll see how to create a `_pushOperation` function to batch the operations.</span></span> <span data-ttu-id="a5d58-128">Сначала посмотрим на следующий пример кода, где показан вызов `_pushOperation` из пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="a5d58-128">First, take a look at the following code example to see how to call `_pushOperation` from your custom function.</span></span>

<span data-ttu-id="a5d58-129">В следующем примере пользовательская функция выполняет деление, обращаясь для этой операции к удаленной службе.</span><span class="sxs-lookup"><span data-stu-id="a5d58-129">In the following code, the custom function performs division but relies on a remote service to do the actual calculation.</span></span> <span data-ttu-id="a5d58-130">Она вызывает `_pushOperation` для включения операции вместе с другими операциями в пакет для удаленной службы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-130">It calls `_pushOperation` to batch the operation along with other operations to the remote service.</span></span> <span data-ttu-id="a5d58-131">Операция здесь называется **div2**.</span><span class="sxs-lookup"><span data-stu-id="a5d58-131">It names the operation **div2**.</span></span> <span data-ttu-id="a5d58-132">Можно использовать для операций любую схему именования, если только в удаленной службе используется такая же схема (дополнительно об удаленной службе см. далее).</span><span class="sxs-lookup"><span data-stu-id="a5d58-132">You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later).</span></span> <span data-ttu-id="a5d58-133">Кроме того передаются аргументы, необходимые удаленной службе для выполнения операции.</span><span class="sxs-lookup"><span data-stu-id="a5d58-133">Also, the arguments the remote service will need to run the operation are passed.</span></span>

### <a name="add-the-div2-custom-function-to-functionsts"></a><span data-ttu-id="a5d58-134">Добавление пользовательской функции div2 в functions.ts</span><span class="sxs-lookup"><span data-stu-id="a5d58-134">Add the div2 custom function to functions.ts</span></span>

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}

CustomFunctions.associate("DIV2", div2);
```

<span data-ttu-id="a5d58-135">После этого следует определить пакетный массив, в котором будут храниться все операции, предназначенные для передачи в одном сетевом вызове.</span><span class="sxs-lookup"><span data-stu-id="a5d58-135">Next, you will define the batch array which will store all operations to be passed in one network call.</span></span> <span data-ttu-id="a5d58-136">В приведенном ниже коде показано, как определить интерфейс, описывающий каждый элемент пакета в массиве.</span><span class="sxs-lookup"><span data-stu-id="a5d58-136">The following code shows how to define an interface describing each batch entry in the array.</span></span> <span data-ttu-id="a5d58-137">Интерфейс определяет операцию, которая представляет собой строку-имя запускаемой операции.</span><span class="sxs-lookup"><span data-stu-id="a5d58-137">The interface defines an operation, which is a string name of which operation to run.</span></span> <span data-ttu-id="a5d58-138">Например, если у вас две пользовательские функции с именами `multiply` и `divide`, их можно использовать как имена операции в элементах пакета.</span><span class="sxs-lookup"><span data-stu-id="a5d58-138">For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries.</span></span> <span data-ttu-id="a5d58-139">`args` будет содержать аргументы, переданные в пользовательскую функцию из Excel.</span><span class="sxs-lookup"><span data-stu-id="a5d58-139">`args` will hold the arguments that were passed to your custom function from Excel.</span></span> <span data-ttu-id="a5d58-140">И, наконец, в `resolve` или `reject` будет храниться обещание с информацией, возвращаемой удаленной службой.</span><span class="sxs-lookup"><span data-stu-id="a5d58-140">And finally, `resolve` or `reject` will store a promise holding the information the remote service returns.</span></span>

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

<span data-ttu-id="a5d58-141">Далее мы создадим пакетный массив, использующий предыдущий интерфейс.</span><span class="sxs-lookup"><span data-stu-id="a5d58-141">Next, create the batch array that uses the previous interface.</span></span> <span data-ttu-id="a5d58-142">Чтобы знать, является ли пакет плановым или нет, создадим переменную _isBatchedRequestSchedule.</span><span class="sxs-lookup"><span data-stu-id="a5d58-142">To track if a batch is scheduled or not, create an \`_isBatchedRequestSchedule variable.</span></span>  <span data-ttu-id="a5d58-143">Она понадобится позже для планирования пакетных вызовов удаленной службы.</span><span class="sxs-lookup"><span data-stu-id="a5d58-143">This will be important later for timing batch calls to the remote service.</span></span>

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

<span data-ttu-id="a5d58-144">Наконец, когда Excel вызывает пользовательскую функцию, необходимо отправить операцию в пакетный массив.</span><span class="sxs-lookup"><span data-stu-id="a5d58-144">Finally when Excel calls your custom function, you need to push the operation into the batch array.</span></span> <span data-ttu-id="a5d58-145">В следующем коде показано, как добавить новую операцию из пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="a5d58-145">The following code shows how to add a new operation from a custom function.</span></span> <span data-ttu-id="a5d58-146">Здесь создается новый элемент пакета, новое обещание для выполнения или отклонения операции, и элемент вставляется в пакетный массив.</span><span class="sxs-lookup"><span data-stu-id="a5d58-146">It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.</span></span>

<span data-ttu-id="a5d58-147">В данном коде также проверяется, является ли пакет плановым.</span><span class="sxs-lookup"><span data-stu-id="a5d58-147">This code also checks to see if a batch is scheduled.</span></span> <span data-ttu-id="a5d58-148">В этом примере выполнение пакете планируется каждые 100 мс.</span><span class="sxs-lookup"><span data-stu-id="a5d58-148">In this example, each batch is scheduled to run every 100ms.</span></span> <span data-ttu-id="a5d58-149">При необходимости этот интервал можно изменить.</span><span class="sxs-lookup"><span data-stu-id="a5d58-149">You can adjust this value as needed.</span></span> <span data-ttu-id="a5d58-150">Чем значение выше, тем больше размер пакета, отправляемого в удаленную службу, и тем дольше пользователь должен ждать результатов.</span><span class="sxs-lookup"><span data-stu-id="a5d58-150">Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results.</span></span> <span data-ttu-id="a5d58-151">При низком значении в удаленную службу отправляется больше пакетов, но зато время ожидания снижается.</span><span class="sxs-lookup"><span data-stu-id="a5d58-151">Lower values tend to send more batches to the remote service, but with a quick response time for users.</span></span>

### <a name="add-the-pushoperation-function-to-functionsts"></a><span data-ttu-id="a5d58-152">Добавление функции `_pushOperation` в functions.ts</span><span class="sxs-lookup"><span data-stu-id="a5d58-152">Add the `_pushOperation` function to functions.ts</span></span>

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a><span data-ttu-id="a5d58-153">Проведение удаленного запроса</span><span class="sxs-lookup"><span data-stu-id="a5d58-153">Make the remote request</span></span>

<span data-ttu-id="a5d58-154">Цель функции `_makeRemoteRequest` – передать пакет операций в удаленную службу, а затем возвратить результаты в каждую пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="a5d58-154">The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function.</span></span> <span data-ttu-id="a5d58-155">Сначала она создает копию пакетного массива.</span><span class="sxs-lookup"><span data-stu-id="a5d58-155">It first creates a copy of the batch array.</span></span> <span data-ttu-id="a5d58-156">Это позволит сразу же начинать включение параллельных вызовов пользовательской функции из Excel в новый массив.</span><span class="sxs-lookup"><span data-stu-id="a5d58-156">This allows concurrent custom function calls from Excel to immediately begin batching in a new array.</span></span> <span data-ttu-id="a5d58-157">Затем копия преобразуется в более простой массив, который не содержит информацию обещания.</span><span class="sxs-lookup"><span data-stu-id="a5d58-157">The copy is then turned into a simpler array that does not contain the promise information.</span></span> <span data-ttu-id="a5d58-158">Не имеет смысла передавать обещания в удаленную службу, так как они не будут работать.</span><span class="sxs-lookup"><span data-stu-id="a5d58-158">It wouldn't make sense to pass the promises to a remote service since they would not work.</span></span> <span data-ttu-id="a5d58-159">Метод \`_makeRemoteRequest будет отклонять или выполнять каждое обещание в зависимости от того, что возвратит удаленная служба.</span><span class="sxs-lookup"><span data-stu-id="a5d58-159">The \`_makeRemoteRequest will either reject or resolve each promise based on what the remote service returns.</span></span>

### <a name="add-the-following-makeremoterequest-method-to-functionsts"></a><span data-ttu-id="a5d58-160">Добавление следующего метода `_makeRemoteRequest` в functions.ts</span><span class="sxs-lookup"><span data-stu-id="a5d58-160">Add the following method to the `_makeRemoteRequest`.</span></span>

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-makeremoterequest-for-your-own-solution"></a><span data-ttu-id="a5d58-161">Переделка `_makeRemoteRequest` для вашего собственного решения</span><span class="sxs-lookup"><span data-stu-id="a5d58-161">Modify `_makeRemoteRequest` for your own solution</span></span>

<span data-ttu-id="a5d58-162">Функция `_makeRemoteRequest` вызывает метод `_fetchFromRemoteService`, который, как будет видно позже, всего лишь имитирует удаленную службу.</span><span class="sxs-lookup"><span data-stu-id="a5d58-162">The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service.</span></span> <span data-ttu-id="a5d58-163">Это упрощает изучение и выполнение кода в данной статье.</span><span class="sxs-lookup"><span data-stu-id="a5d58-163">This makes it easier to study and run the code in this article.</span></span> <span data-ttu-id="a5d58-164">Но если вы хотите использовать этот код для реальной удаленной службы, в него необходимо внести следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="a5d58-164">But when you want to use this code for an actual remote service you should make the following changes:</span></span>

- <span data-ttu-id="a5d58-165">Выберите способ сериализации пакетных операций по сети.</span><span class="sxs-lookup"><span data-stu-id="a5d58-165">Decide how to serialize the batch operations over the network.</span></span> <span data-ttu-id="a5d58-166">Например может потребоваться поместить массива в текст JSON.</span><span class="sxs-lookup"><span data-stu-id="a5d58-166">For example, you may want to put the array into a JSON body.</span></span>
- <span data-ttu-id="a5d58-167">Вместо вызова `_fetchFromRemoteService` следует сделать сетевой вызов удаленной службы с передачей пакета операций.</span><span class="sxs-lookup"><span data-stu-id="a5d58-167">Instead of calling `_fetchFromRemoteService` you'll need to make the actual network call to the remote service passing the batch of operations.</span></span>

## <a name="process-the-batch-call-on-the-remote-service"></a><span data-ttu-id="a5d58-168">Обработка пакетного вызова в удаленной службе</span><span class="sxs-lookup"><span data-stu-id="a5d58-168">Process the batch call on the remote service</span></span>

<span data-ttu-id="a5d58-169">Последний шаг – это выполнение пакетного вызова в удаленной службе.</span><span class="sxs-lookup"><span data-stu-id="a5d58-169">The last step is to handle the batch call in the remote service.</span></span> <span data-ttu-id="a5d58-170">В следующем примере кода показана функция `_fetchFromRemoteService`.</span><span class="sxs-lookup"><span data-stu-id="a5d58-170">The following code sample shows the `_fetchFromRemoteService` function.</span></span> <span data-ttu-id="a5d58-171">Эта функция распаковывает каждую операцию, выполняет указанную операцию и возвращает результат.</span><span class="sxs-lookup"><span data-stu-id="a5d58-171">This function unpacks each operation, performs the specified operation, and returns the results.</span></span> <span data-ttu-id="a5d58-172">Для учебных целей в данной статье применяется функция `_fetchFromRemoteService`, которая запускается в вашей веб-надстройке и имитирует удаленную службу.</span><span class="sxs-lookup"><span data-stu-id="a5d58-172">For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service.</span></span> <span data-ttu-id="a5d58-173">Этот код можно добавить в файл functions.ts, чтобы изучать и запускать его, не создавая настоящую удаленную службу.</span><span class="sxs-lookup"><span data-stu-id="a5d58-173">You can add this code to your functions.ts file so that you can study and run all the code in this article without having to set up an actual remote service.</span></span>

### <a name="add-the-following-fetchfromremoteservice-function-to-functionsts"></a><span data-ttu-id="a5d58-174">Добавление следующей функции `_fetchFromRemoteService` в functions.ts</span><span class="sxs-lookup"><span data-stu-id="a5d58-174">Add the following function to `_fetchFromRemoteService`.</span></span>

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-fetchfromremoteservice-for-your-live-remote-service"></a><span data-ttu-id="a5d58-175">Переделка `_fetchFromRemoteService` для действующей удаленной службы</span><span class="sxs-lookup"><span data-stu-id="a5d58-175">Modify `_fetchFromRemoteService` for your live remote service</span></span>

<span data-ttu-id="a5d58-176">Чтобы переделать функцию `_fetchFromRemoteService` для выполнения в действующей удаленной службе, внесите следующие изменения.</span><span class="sxs-lookup"><span data-stu-id="a5d58-176">To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes:</span></span>

- <span data-ttu-id="a5d58-177">В зависимости от платформы используемого сервера (Node.js или другая) сопоставьте сетевой вызов клиента с этой функцией.</span><span class="sxs-lookup"><span data-stu-id="a5d58-177">Depending on your server platform (Node.js or others) map the client network call to this function.</span></span>
- <span data-ttu-id="a5d58-178">Удалите функцию `pause`, которая имитирует задержку в сети.</span><span class="sxs-lookup"><span data-stu-id="a5d58-178">Remove the `pause` function which simulates network latency as part of the mock.</span></span>
- <span data-ttu-id="a5d58-179">Измените объявление функции так, чтобы она работала с переданным параметром, если параметр изменяется для целей сети.</span><span class="sxs-lookup"><span data-stu-id="a5d58-179">Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes.</span></span> <span data-ttu-id="a5d58-180">Например, это может быть не массив а текст JSON, содержащий требуемые пакетные операции.</span><span class="sxs-lookup"><span data-stu-id="a5d58-180">For example, instead of an array, it may be a JSON body of batched operations to process.</span></span>
- <span data-ttu-id="a5d58-181">Переделайте функцию для выполнения операций (или вызова функций, которые выполняют операции).</span><span class="sxs-lookup"><span data-stu-id="a5d58-181">Modify the function to perform the operations (or call functions that do the operations).</span></span>
- <span data-ttu-id="a5d58-182">Примените подходящий механизм проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a5d58-182">Apply an appropriate authentication mechanism.</span></span> <span data-ttu-id="a5d58-183">Убедитесь, что доступ к функции есть только у предусмотренных вами вызывающих пользователей.</span><span class="sxs-lookup"><span data-stu-id="a5d58-183">Ensure that only the correct callers can access the function.</span></span>
- <span data-ttu-id="a5d58-184">Поместите код в удаленную службу.</span><span class="sxs-lookup"><span data-stu-id="a5d58-184">Place the code in the remote service.</span></span>

## <a name="see-also"></a><span data-ttu-id="a5d58-185">См. также</span><span class="sxs-lookup"><span data-stu-id="a5d58-185">See also</span></span>

* [<span data-ttu-id="a5d58-186">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="a5d58-186">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a5d58-187">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="a5d58-187">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a5d58-188">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="a5d58-188">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="a5d58-189">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="a5d58-189">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)