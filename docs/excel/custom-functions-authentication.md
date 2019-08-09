---
ms.date: 07/09/2019
description: Проверка подлинности пользователей с использованием пользовательских функций в Excel.
title: Проверка подлинности для пользовательских функций
localization_priority: Priority
ms.openlocfilehash: f746947122da7ef3d54a0dd3b4f90dd059e5830f
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268140"
---
# <a name="authentication-for-custom-functions"></a><span data-ttu-id="b770e-103">Проверка подлинности для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="b770e-103">Authentication for custom functions</span></span>

<span data-ttu-id="b770e-104">В некоторых случаях пользовательским функциям требуется проверить подлинность пользователя, чтобы получить доступ к защищенным ресурсам.</span><span class="sxs-lookup"><span data-stu-id="b770e-104">In some scenarios your custom function will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="b770e-105">Хотя пользовательские функции не требуют определенного метода проверки подлинности, следует учитывать, что они выполняются в отдельной среде из области задач и других элементов пользовательского интерфейса надстройки.</span><span class="sxs-lookup"><span data-stu-id="b770e-105">While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in.</span></span> <span data-ttu-id="b770e-106">Поэтому между двумя средами выполнения требуется осуществлять обмен данными с помощью объекта `OfficeRuntime.storage` и API диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="b770e-106">Because of this, you'll need to pass data back and forth between the two runtimes using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="b770e-107">Объект OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="b770e-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="b770e-108">В среде выполнения пользовательских функций отсутствует объект `localStorage`, доступный в глобальном окне, где обычно хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="b770e-108">The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data.</span></span> <span data-ttu-id="b770e-109">Вместо этого следует обмениваться данными между пользовательскими функциями и областями задач, используя объект [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) для настройки и получения данных.</span><span class="sxs-lookup"><span data-stu-id="b770e-109">Instead, you should share data between custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

<span data-ttu-id="b770e-110">Кроме того, использование объекта `storage` дает преимущество; он использует безопасную изолированную среду, чтобы ваши данные были недоступны другим надстройкам.</span><span class="sxs-lookup"><span data-stu-id="b770e-110">Additionally, there is a benefit to using the `storage` object; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="b770e-111">Рекомендуемое использование</span><span class="sxs-lookup"><span data-stu-id="b770e-111">Suggested usage</span></span>

<span data-ttu-id="b770e-112">Если вам нужно проверить подлинность из области задач или пользовательской функции, проверьте объект `storage`, чтобы узнать, был ли уже получен маркер доступа.</span><span class="sxs-lookup"><span data-stu-id="b770e-112">When you need to authenticate either from the task pane or a custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="b770e-113">Если нет, используйте API диалоговых окон, чтобы проверить подлинность пользователя, извлечь маркер доступа и сохранить его в объекте `storage` для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="b770e-113">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="b770e-114">API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="b770e-114">Dialog API example</span></span>

<span data-ttu-id="b770e-115">Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя войти в систему.</span><span class="sxs-lookup"><span data-stu-id="b770e-115">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="b770e-116">После ввода пользователем своих учетных данных итоговый маркер доступа можно сохранить в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-116">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="b770e-117">В среде выполнения пользовательских функций используется объект Dialog, немного отличающийся от объекта Dialog среды выполнения модуля браузера, применяемого областями задач.</span><span class="sxs-lookup"><span data-stu-id="b770e-117">The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="b770e-118">Они оба называются "API диалоговых окон", но для проверки подлинности пользователей в среде выполнения пользовательских функций используется интерфейс `OfficeRuntime.Dialog`.</span><span class="sxs-lookup"><span data-stu-id="b770e-118">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the custom functions runtime.</span></span>

<span data-ttu-id="b770e-119">Сведения об использовании объекта `Dialog` см. в статье [Диалоговое окно пользовательских функций](/office/dev/add-ins/excel/custom-functions-dialog).</span><span class="sxs-lookup"><span data-stu-id="b770e-119">For information on how to use the `Dialog` object, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).</span></span>

<span data-ttu-id="b770e-120">При формировании всего процесса проверки подлинности рекомендуется рассматривать область задач, элементы пользовательского интерфейса надстройки и часть надстройки, включающую пользовательские функции, как отдельные объекты, которые могут взаимодействовать друг с другом с помощью интерфейса `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-120">When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `OfficeRuntime.storage`.</span></span>

<span data-ttu-id="b770e-121">На следующей схеме показан этот основной процесс.</span><span class="sxs-lookup"><span data-stu-id="b770e-121">The following diagram outlines this basic process.</span></span> <span data-ttu-id="b770e-122">Обратите внимание на пунктирную линию, которая указывает, что хотя пользовательские функции выполняют отдельные действия, они и область задач надстройки входят в состав надстройки.</span><span class="sxs-lookup"><span data-stu-id="b770e-122">Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.</span></span>

1. <span data-ttu-id="b770e-123">Вы выполняете вызов пользовательской функции из ячейки книги Excel.</span><span class="sxs-lookup"><span data-stu-id="b770e-123">You issue a custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="b770e-124">Пользовательская функция использует `Dialog` для передачи учетных данных пользователя на веб-сайт.</span><span class="sxs-lookup"><span data-stu-id="b770e-124">The custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="b770e-125">Этот веб-сайт возвращает маркер доступа для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="b770e-125">This website then returns an access token to the custom function.</span></span>
4. <span data-ttu-id="b770e-126">После этого пользовательская функция устанавливает этот маркер доступа в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-126">Your custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="b770e-127">Область задач надстройки получает доступ к маркеру из объекта `storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-127">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="b770e-128">![Схема пользовательской функции, использующей API диалоговых окон для получения маркера доступа, с последующим предоставлением маркера для области задач с помощью API OfficeRuntime.storage. ](../images/authentication-diagram.png "Схема проверки подлинности.")</span><span class="sxs-lookup"><span data-stu-id="b770e-128">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="b770e-129">Хранение маркера</span><span class="sxs-lookup"><span data-stu-id="b770e-129">Storing the token</span></span>

<span data-ttu-id="b770e-130">Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage).</span><span class="sxs-lookup"><span data-stu-id="b770e-130">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="b770e-131">Этот пример кода представляет полный пример обмена данными между пользовательскими функциями и областью задач.</span><span class="sxs-lookup"><span data-stu-id="b770e-131">Refer to this code sample for a complete example of sharing data between custom functions and the task pane.</span></span>

<span data-ttu-id="b770e-132">При проверке подлинности пользовательской функцией она получает маркер доступа, который должен храниться в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-132">If the custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="b770e-133">В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение.</span><span class="sxs-lookup"><span data-stu-id="b770e-133">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="b770e-134">Функция `storeValue` — это пользовательская функция, которая в данном примере сохраняет значение, полученное от пользователя.</span><span class="sxs-lookup"><span data-stu-id="b770e-134">The `storeValue` function is a custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="b770e-135">Можно внести изменение, чтобы сохранять любые нужные значения маркеров.</span><span class="sxs-lookup"><span data-stu-id="b770e-135">You can modify this to store any token value you need.</span></span>

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

<span data-ttu-id="b770e-136">Когда области задач требуется маркер доступа, она может извлечь его из объекта `storage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-136">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="b770e-137">В следующем примере кода показано, как использовать метод `storage.getItem` чтобы извлечь маркер.</span><span class="sxs-lookup"><span data-stu-id="b770e-137">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a><span data-ttu-id="b770e-138">Общие рекомендации</span><span class="sxs-lookup"><span data-stu-id="b770e-138">General Guidance</span></span>

<span data-ttu-id="b770e-139">Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="b770e-139">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="b770e-140">При реализации своей собственной проверки подлинности с использованием пользовательских функций отсутствует определенный шаблон или метод.</span><span class="sxs-lookup"><span data-stu-id="b770e-140">There is no particular pattern or method you must follow to implement your own authentication with custom functions.</span></span> <span data-ttu-id="b770e-141">Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="b770e-141">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins).</span></span>  

<span data-ttu-id="b770e-142">Избегайте использования следующих расположений для хранения данных при разработке пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="b770e-142">Avoid using the following locations to store data when developing custom functions:</span></span>  

- <span data-ttu-id="b770e-143">`localStorage`: у пользовательских функций нет доступа к глобальному объекту `window`, поэтому им недоступны данные, хранящиеся в объекте `localStorage`.</span><span class="sxs-lookup"><span data-stu-id="b770e-143">`localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="b770e-144">`Office.context.document.settings`: это расположение не защищено, и сведения могут быть извлечены любым пользователем с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="b770e-144">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b770e-145">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="b770e-145">Next steps</span></span>
<span data-ttu-id="b770e-146">Сведения об [API диалоговых окон для пользовательских функций](custom-functions-dialog.md).</span><span class="sxs-lookup"><span data-stu-id="b770e-146">Learn about the [dialog API for custom functions](custom-functions-dialog.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b770e-147">См. также</span><span class="sxs-lookup"><span data-stu-id="b770e-147">See also</span></span>

* [<span data-ttu-id="b770e-148">Архитектура пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="b770e-148">Custom functions architecture</span></span>](custom-functions-architecture.md)
* [<span data-ttu-id="b770e-149">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="b770e-149">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="b770e-150">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="b770e-150">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="b770e-151">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="b770e-151">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
