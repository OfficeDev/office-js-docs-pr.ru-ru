---
title: Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы
description: Узнайте обходные пути, которые можно использовать, если метод messageChild не поддерживается.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8da6bc3e1231bc6296a16fa153dc0e4ba1bd102b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349779"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="1846f-103">Альтернативные способы передачи сообщений в диалоговое окно со своей хост-страницы</span><span class="sxs-lookup"><span data-stu-id="1846f-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="1846f-104">Рекомендуемый способ передачи данных и сообщений с родительской страницы в диалоговое окно для детей используется метод, описанный в API диалоговых Office в Office надстройки `messageChild` . [](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) Если надстройка работает на платформе или хосте, не поддерживающей набор требований [DialogApi 1.2,](../reference/requirement-sets/dialog-api-requirement-sets.md)существует два других способа передать информацию в диалоговое окно:</span><span class="sxs-lookup"><span data-stu-id="1846f-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="1846f-105">Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="1846f-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="1846f-106">Храните информацию в месте, доступном как для главного, так и для диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="1846f-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="1846f-107">Два окна не имеют общего хранилища сеансов (свойство [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) но если они имеют один и тот же домен *(включая* номер [порта,](https://www.w3schools.com/html/html5_webstorage.asp)если таковые имеются), они имеют общий локальный служба хранилища .\*</span><span class="sxs-lookup"><span data-stu-id="1846f-107">The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="1846f-108">\* Существует ошибка, влияющая на вашу стратегию обработки маркеров.</span><span class="sxs-lookup"><span data-stu-id="1846f-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="1846f-109">Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.</span><span class="sxs-lookup"><span data-stu-id="1846f-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="1846f-110">Использование локального хранилища</span><span class="sxs-lookup"><span data-stu-id="1846f-110">Use local storage</span></span>

<span data-ttu-id="1846f-111">Чтобы использовать локальное хранилище, перед вызовом вызываем метод объекта на хост-странице, `setItem` `window.localStorage` как в следующем `displayDialogAsync` примере.</span><span class="sxs-lookup"><span data-stu-id="1846f-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example.</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="1846f-112">Код в диалоговом окне читает элемент при необходимости, как в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="1846f-112">Code in the dialog box reads the item when it's needed, as in the following example.</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="1846f-113">Использование параметров запроса</span><span class="sxs-lookup"><span data-stu-id="1846f-113">Use query parameters</span></span>

<span data-ttu-id="1846f-114">В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.</span><span class="sxs-lookup"><span data-stu-id="1846f-114">The following example shows how to pass data with a query parameter.</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="1846f-115">Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="1846f-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="1846f-116">Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.</span><span class="sxs-lookup"><span data-stu-id="1846f-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1846f-117">Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="1846f-117">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="1846f-118">(Этот параметр добавляется после пользовательских параметров запроса, если они есть.</span><span class="sxs-lookup"><span data-stu-id="1846f-118">(It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="1846f-119">Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать.</span><span class="sxs-lookup"><span data-stu-id="1846f-119">It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it.</span></span> <span data-ttu-id="1846f-120">Это же значение добавляется в хранилище сеансов диалоговое окно (свойство [Window.sessionStorage).](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)</span><span class="sxs-lookup"><span data-stu-id="1846f-120">The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="1846f-121">*Ваш код не должен ни считывать это значение, ни записывать в него данные*.</span><span class="sxs-lookup"><span data-stu-id="1846f-121">Again, *your code should neither read nor write to this value*.</span></span>
