---
ms.date: 05/17/2020
description: Проверка подлинности пользователей с помощью настраиваемой Excel, которые не используют области задач.
title: Проверка подлинности для пользовательских функций без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 94eadd343f969e6dbd83881764fac936acf0704b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349702"
---
# <a name="authentication-for-ui-less-custom-functions"></a><span data-ttu-id="3d495-103">Проверка подлинности для пользовательских функций без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="3d495-103">Authentication for UI-less custom functions</span></span>

<span data-ttu-id="3d495-104">В некоторых сценариях вашей настраиваемой функции, которая не использует области задач или других элементов пользовательского интерфейса (настраиваемая функция без пользовательского интерфейса), потребуется проверить подлинность пользователя, чтобы получить доступ к защищенным ресурсам.</span><span class="sxs-lookup"><span data-stu-id="3d495-104">In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources.</span></span> <span data-ttu-id="3d495-105">Следует помнить, что пользовательские функции без пользовательского интерфейса выполняются только для JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3d495-105">Be aware that UI-less custom functions run in a JavaScript-only runtime.</span></span> <span data-ttu-id="3d495-106">Из-за этого необходимо передавать данные между временем запуска только JavaScript и типичным временем запуска браузера, используемым большинством надстройок с помощью объекта и `OfficeRuntime.storage` API диалогов.</span><span class="sxs-lookup"><span data-stu-id="3d495-106">Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a><span data-ttu-id="3d495-107">Объект OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="3d495-107">OfficeRuntime.storage object</span></span>

<span data-ttu-id="3d495-108">Время запуска только на JavaScript, используемое пользовательскими функциями без пользовательского интерфейса, не имеет объекта, доступного в глобальном окне, где обычно `localStorage` хранятся данные.</span><span class="sxs-lookup"><span data-stu-id="3d495-108">The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data.</span></span> <span data-ttu-id="3d495-109">Вместо этого следует обмениваться данными между пользовательскими функциями и области задач без пользовательского интерфейса, используя [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) для настройки и получения данных.</span><span class="sxs-lookup"><span data-stu-id="3d495-109">Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.</span></span>

### <a name="suggested-usage"></a><span data-ttu-id="3d495-110">Рекомендуемое использование</span><span class="sxs-lookup"><span data-stu-id="3d495-110">Suggested usage</span></span>

<span data-ttu-id="3d495-111">При необходимости проверки подлинности из настраиваемой функции, не входя в пользовательский интерфейс, проверьте, был ли уже приобретен маркер `storage` доступа.</span><span class="sxs-lookup"><span data-stu-id="3d495-111">When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired.</span></span> <span data-ttu-id="3d495-112">Если нет, используйте API диалоговых окон, чтобы проверить подлинность пользователя, извлечь маркер доступа и сохранить его в объекте `storage` для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="3d495-112">If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.</span></span>

## <a name="dialog-api"></a><span data-ttu-id="3d495-113">API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="3d495-113">Dialog API</span></span>

<span data-ttu-id="3d495-114">Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя войти в систему.</span><span class="sxs-lookup"><span data-stu-id="3d495-114">If a token doesn't exist, you should use the Dialog API to ask the user to sign in.</span></span> <span data-ttu-id="3d495-115">После ввода пользователем своих учетных данных итоговый маркер доступа можно сохранить в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="3d495-115">After a user enters their credentials, the resulting access token can be stored in `storage`.</span></span>

> [!NOTE]
> <span data-ttu-id="3d495-116">Время запуска только на JavaScript использует объект Dialog, который немного отличается от объекта Dialog в времени запуска двигателя браузера, используемого в области задач.</span><span class="sxs-lookup"><span data-stu-id="3d495-116">The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes.</span></span> <span data-ttu-id="3d495-117">Они оба называются "API диалогов", но используются для проверки подлинности пользователей в время запуска только `OfficeRuntime.Dialog` javaScript.</span><span class="sxs-lookup"><span data-stu-id="3d495-117">They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.</span></span>

<span data-ttu-id="3d495-118">На следующей схеме показан этот основной процесс.</span><span class="sxs-lookup"><span data-stu-id="3d495-118">The following diagram outlines this basic process.</span></span> <span data-ttu-id="3d495-119">Пунктирная строка указывает на то, что пользовательские функции без пользовательского интерфейса и области задач надстройки являются частью надстройки в целом, хотя они используют отдельные время запуска.</span><span class="sxs-lookup"><span data-stu-id="3d495-119">The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.</span></span>

1. <span data-ttu-id="3d495-120">Вы выдает пользовательский вызов функции без пользовательского интерфейса из ячейки в Excel книге.</span><span class="sxs-lookup"><span data-stu-id="3d495-120">You issue a UI-less custom function call from a cell in an Excel workbook.</span></span>
2. <span data-ttu-id="3d495-121">Настраиваемая функция, не использующая пользовательский интерфейс, используется для того, чтобы передать учетные данные пользователя `Dialog` веб-сайту.</span><span class="sxs-lookup"><span data-stu-id="3d495-121">The UI-less custom function uses `Dialog` to pass your user credentials to a website.</span></span>
3. <span data-ttu-id="3d495-122">Затем этот веб-сайт возвращает маркер доступа к настраиваемой функции без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="3d495-122">This website then returns an access token to the UI-less custom function.</span></span>
4. <span data-ttu-id="3d495-123">Настраиваемая функция без пользовательского интерфейса задает этот маркер доступа к `storage` .</span><span class="sxs-lookup"><span data-stu-id="3d495-123">Your UI-less custom function then sets this access token to the `storage`.</span></span>
5. <span data-ttu-id="3d495-124">Область задач надстройки получает доступ к маркеру из объекта `storage`.</span><span class="sxs-lookup"><span data-stu-id="3d495-124">Your add-in's task pane accesses the token from `storage`.</span></span>

<span data-ttu-id="3d495-125">![Схема настраиваемой функции с помощью диалогового API для получения маркера доступа, а затем обмена маркером с области задач через API OfficeRuntime.storage.](../images/authentication-diagram.png "Схема проверки подлинности.")</span><span class="sxs-lookup"><span data-stu-id="3d495-125">![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")</span></span>

## <a name="storing-the-token"></a><span data-ttu-id="3d495-126">Хранение маркера</span><span class="sxs-lookup"><span data-stu-id="3d495-126">Storing the token</span></span>

<span data-ttu-id="3d495-127">Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage).</span><span class="sxs-lookup"><span data-stu-id="3d495-127">The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample.</span></span> <span data-ttu-id="3d495-128">Обратитесь к этому примеру кода для полного примера обмена данными между пользовательскими функциями без пользовательского интерфейса и области задач.</span><span class="sxs-lookup"><span data-stu-id="3d495-128">Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.</span></span>

<span data-ttu-id="3d495-129">Если настраиваемая функция без пользовательского интерфейса подает проверку подлинности, она получает маркер доступа и должна будет хранить его `storage` в .</span><span class="sxs-lookup"><span data-stu-id="3d495-129">If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`.</span></span> <span data-ttu-id="3d495-130">В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение.</span><span class="sxs-lookup"><span data-stu-id="3d495-130">The following code sample shows how to call the `storage.setItem` method to store a value.</span></span> <span data-ttu-id="3d495-131">Функция — это настраиваемая функция без пользовательского интерфейса, которая, например, `storeValue` сохраняет значение от пользователя.</span><span class="sxs-lookup"><span data-stu-id="3d495-131">The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user.</span></span> <span data-ttu-id="3d495-132">Можно внести изменение, чтобы сохранять любые нужные значения маркеров.</span><span class="sxs-lookup"><span data-stu-id="3d495-132">You can modify this to store any token value you need.</span></span>

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

<span data-ttu-id="3d495-133">Когда области задач требуется маркер доступа, она может извлечь его из объекта `storage`.</span><span class="sxs-lookup"><span data-stu-id="3d495-133">When the task pane needs the access token, it can retrieve the token from `storage`.</span></span> <span data-ttu-id="3d495-134">В следующем примере кода показано, как использовать метод `storage.getItem` чтобы извлечь маркер.</span><span class="sxs-lookup"><span data-stu-id="3d495-134">The following code sample shows how to use the `storage.getItem` method to retrieve the token.</span></span>

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

## <a name="general-guidance"></a><span data-ttu-id="3d495-135">Общие рекомендации</span><span class="sxs-lookup"><span data-stu-id="3d495-135">General guidance</span></span>

<span data-ttu-id="3d495-136">Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="3d495-136">Office Add-ins are web-based and you can use any web authentication technique.</span></span> <span data-ttu-id="3d495-137">Для реализации собственной проверки подлинности с помощью пользовательских функций не существует определенного шаблона или метода, который необходимо выполнить.</span><span class="sxs-lookup"><span data-stu-id="3d495-137">There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions.</span></span> <span data-ttu-id="3d495-138">Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](../develop/auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="3d495-138">You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).</span></span>  

<span data-ttu-id="3d495-139">Избегайте использования следующих местоположений для хранения данных при разработке настраиваемой функции: .</span><span class="sxs-lookup"><span data-stu-id="3d495-139">Avoid using the following locations to store data when developing custom functions: .</span></span>

- <span data-ttu-id="3d495-140">`localStorage`. Пользовательские функции, не влияемые на пользовательский интерфейс, не имеют доступа к глобальному объекту и поэтому не имеют доступа к данным, `window` хранимым в `localStorage` .</span><span class="sxs-lookup"><span data-stu-id="3d495-140">`localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.</span></span>
- <span data-ttu-id="3d495-141">`Office.context.document.settings`: это расположение не защищено, и сведения могут быть извлечены любым пользователем с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="3d495-141">`Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="3d495-142">Пример API диалоговых полей</span><span class="sxs-lookup"><span data-stu-id="3d495-142">Dialog box API example</span></span>

<span data-ttu-id="3d495-143">В следующем примере кода функция использует функцию API для отображения `getTokenViaDialog` `Dialog` `displayWebDialogOptions` диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="3d495-143">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span> <span data-ttu-id="3d495-144">Этот пример предоставляется для демонстрации возможностей объекта, а не для проверки `Dialog` подлинности.</span><span class="sxs-lookup"><span data-stu-id="3d495-144">This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.</span></span>

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a><span data-ttu-id="3d495-145">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="3d495-145">Next steps</span></span>
<span data-ttu-id="3d495-146">Узнайте, как [отламыть пользовательские функции без пользовательского интерфейса.](custom-functions-debugging.md)</span><span class="sxs-lookup"><span data-stu-id="3d495-146">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3d495-147">См. также</span><span class="sxs-lookup"><span data-stu-id="3d495-147">See also</span></span>

* [<span data-ttu-id="3d495-148">Время запуска для пользовательских Excel пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="3d495-148">Runtime for UI-less Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3d495-149">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="3d495-149">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)