---
ms.date: 03/19/2019
description: Проверка поДлинности пользователей с помощью пользовательских функций в Excel.
title: Проверка поДлинности для пользовательских функций
ms.openlocfilehash: 7db46e40758ea0282a2fd7c4d40739304a874e76
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871495"
---
# <a name="authentication"></a><span data-ttu-id="85231-103">Проверка подлинности</span><span class="sxs-lookup"><span data-stu-id="85231-103">Authentication</span></span>

<span data-ttu-id="85231-104">В некоторых сценариях пользовательская функция должна проверить подлинность пользователя, чтобы получить доступ к защищенным ресурсам.</span><span class="sxs-lookup"><span data-stu-id="85231-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="85231-105">В то время как для пользовательских функций не требуется определенный способ проверки подлинности, следует учитывать, что пользовательские функции выполняются в отдельной среде выполнения из области задач и других элементов ПОЛЬЗОВАТЕЛЬСКОГО интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="85231-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="85231-106">Поэтому необходимо передавать данные между двумя средами выполнения с помощью `AsyncStorage` объекта и API диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="85231-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.</span></span>
  
## <a name="asyncstorage-object"></a><span data-ttu-id="85231-107">Объект Асинкстораже</span><span class="sxs-lookup"><span data-stu-id="85231-107">AsyncStorage object</span></span>

<span data-ttu-id="85231-108">В среде выполнения пользовательских функций отсутствует `localStorage` объект, доступный в глобальном окне, где обычно могут храниться данные.</span><span class="sxs-lookup"><span data-stu-id="85231-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="85231-109">Вместо этого следует обмениваться данными между пользовательскими функциями и областями задач с помощью [оффицерунтиме. асинкстораже](/javascript/api/office-runtime/officeruntime.asyncstorage) для задания и получения данных.</span><span class="sxs-lookup"><span data-stu-id="85231-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.</span></span>

<span data-ttu-id="85231-110">Кроме того, существует преимущество использования `AsyncStorage`; Она использует безопасную изолированную среду, чтобы получить доступ к данным другими надстройками.</span><span class="sxs-lookup"><span data-stu-id="85231-110">Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="85231-111">Предлагаемое использование</span><span class="sxs-lookup"><span data-stu-id="85231-111">Suggested usage</span></span>

<span data-ttu-id="85231-112">Если вам нужно выполнить проверку подлинности из области задач или настраиваемой функции, `AsyncStorage` проверьте, был ли уже получен маркер доступа.</span><span class="sxs-lookup"><span data-stu-id="85231-112">When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired.</span></span> <span data-ttu-id="85231-113">В противном случае используйте API диалоговых окон для проверки подлинности пользователя, получения маркера доступа и сохранения маркера в `AsyncStorage` для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="85231-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="85231-114">API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="85231-114">Dialog API</span></span>

<span data-ttu-id="85231-115">Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="85231-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="85231-116">После того как пользователь введет свои учетные данные, получившийся маркер доступа можно `AsyncStorage`будет хранить в файле.</span><span class="sxs-lookup"><span data-stu-id="85231-116">After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.</span></span>

> [!NOTE]
> <span data-ttu-id="85231-117">В среде выполнения пользовательских функций используется объект Dialog, который немного отличается от объекта Dialog в среде выполнения модуля браузера, используемого панелями задач.</span><span class="sxs-lookup"><span data-stu-id="85231-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="85231-118">Они обе называются "диалоговым API", но используются `Officeruntime.Dialog` для проверки подлинности пользователей в среде выполнения пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="85231-118">They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="85231-119">Сведения о том `OfficeRuntime.Dialog`, как использовать, можно найти в разделе [Среда выполнения пользовательских функций](/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span><span class="sxs-lookup"><span data-stu-id="85231-119">For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).</span></span>

<span data-ttu-id="85231-120">Выполняя весь процесс проверки подлинности в целом, можно представить себе область задач и элементы ПОЛЬЗОВАТЕЛЬСКОГО интерфейса надстройки, а также компонент Custom functions в надстройке как отдельные объекты, которые могут общаться друг с другом с помощью `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="85231-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `AsyncStorage`.</span></span>

<span data-ttu-id="85231-121">Ниже приведена схема, в которой описан этот базовый процесс.</span><span class="sxs-lookup"><span data-stu-id="85231-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="85231-122">Обратите внимание, что пунктирная линия указывает на то, что при выполнении отдельных действий пользовательские функции и область задач надстройки являются частью надстройки в целом.</span><span class="sxs-lookup"><span data-stu-id="85231-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="85231-123">Вы выДаете вызов пользовательской функции из ячейки книги Excel.</span><span class="sxs-lookup"><span data-stu-id="85231-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="85231-124">Пользовательская функция использует `Officeruntime.Dialog` для передачи учетных данных пользователя на веб-сайт.</span><span class="sxs-lookup"><span data-stu-id="85231-124">The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="85231-125">Затем этот веб-сайт возвращает маркер доступа к пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="85231-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="85231-126">Затем пользовательская функция устанавливает для маркера доступа значение `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="85231-126">Your custom function then sets this access token to the `AsyncStorage`.</span></span>
5. <span data-ttu-id="85231-127">Область задач надстройки получает доступ к маркеру из `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="85231-127">Your add-in's task pane accesses the token from `AsyncStorage`.</span></span>

<span data-ttu-id="85231-128">![Схема пользовательской функции с помощью API диалога для получения маркера доступа, а затем совместного использования маркера с областью задач с помощью API асинкстораже.] (../images/authentication-diagram.png "Схема проверки подлинности.")</span><span class="sxs-lookup"><span data-stu-id="85231-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the AsyncStorage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="85231-129">Сохранение маркера</span><span class="sxs-lookup"><span data-stu-id="85231-129">Storing the token</span></span>

<span data-ttu-id="85231-130">Приведенные ниже примеры относятся к [использованию примера использования асинкстораже в коде пользовательских функций](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) .</span><span class="sxs-lookup"><span data-stu-id="85231-130">The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="85231-131">В этом примере кода приведен полный пример совместного использования данных пользовательскими функциями и областью задач.</span><span class="sxs-lookup"><span data-stu-id="85231-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="85231-132">Если пользовательская функция выполняет проверку подлинности, она получает маркер доступа и должна храниться в `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="85231-132">If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`.</span></span> <span data-ttu-id="85231-133">В приведенном ниже примере кода показано, как `AsyncStorage.setItem` вызвать метод для хранения значения.</span><span class="sxs-lookup"><span data-stu-id="85231-133">The following code sample shows how to call the `AsyncStorage.setItem` method to store a value.</span></span> <span data-ttu-id="85231-134">`StoreValue` Функция — это пользовательская функция, в которой пример содержит значение от пользователя.</span><span class="sxs-lookup"><span data-stu-id="85231-134">The `StoreValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="85231-135">Вы можете изменить его, чтобы сохранить все необходимые значения маркера.</span><span class="sxs-lookup"><span data-stu-id="85231-135">You can modify this to store any token value you need.</span></span>

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

<span data-ttu-id="85231-136">Когда область задач нуждается в маркере доступа, она может получить маркер из `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="85231-136">When the task pane needs the access token, it can retrieve the token from `AsyncStorage`.</span></span> <span data-ttu-id="85231-137">В приведенном ниже примере кода показано, как `AsyncStorage.getItem` использовать метод для получения маркера.</span><span class="sxs-lookup"><span data-stu-id="85231-137">The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.</span></span>

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## <a name="general-guidance"></a><span data-ttu-id="85231-138">Общие рекомендации</span><span class="sxs-lookup"><span data-stu-id="85231-138">General guidance</span></span>

<span data-ttu-id="85231-139">Надстройки Office основаны на веб-интерфейсе, и вы можете использовать любой способ проверки подлинности веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="85231-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="85231-140">Для реализации собственной проверки подлинности с настраиваемыми функциями не нужно выполнять определенные шаблоны или методы.</span><span class="sxs-lookup"><span data-stu-id="85231-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="85231-141">Вы можете ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи, посвященной авторизации через внешние службы](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="85231-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).</span></span>  

<span data-ttu-id="85231-142">Избегайте использования следующих расположений для хранения данных при разработке пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="85231-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="85231-143">`localStorage`: Пользовательские функции не имеют доступа к глобальному `window` объекту и поэтому не имеют доступа к данным, хранящимся `localStorage`в.</span><span class="sxs-lookup"><span data-stu-id="85231-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.</span></span>
- <span data-ttu-id="85231-144">`Office.context.document.settings`: Это расположение не является безопасным, и его данные могут извлекаться кем угодно с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="85231-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="85231-145">См. также</span><span class="sxs-lookup"><span data-stu-id="85231-145">See also</span></span>

* [<span data-ttu-id="85231-146">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="85231-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="85231-147">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="85231-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="85231-148">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="85231-148">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="85231-149">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="85231-149">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
