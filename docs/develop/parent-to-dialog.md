---
title: Альтернативные способы передачи сообщений в диалоговое окно с главной страницы
description: Узнайте, как использовать методы обхода, если метод Мессажечилд не поддерживается.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8f44f7f5c145b58d13e7387d01e28fd349a512fc
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279486"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a><span data-ttu-id="3b686-103">Альтернативные способы передачи сообщений в диалоговое окно с главной страницы</span><span class="sxs-lookup"><span data-stu-id="3b686-103">Alternative ways of passing messages to a dialog box from its host page</span></span>

<span data-ttu-id="3b686-104">Рекомендуемый способ передачи данных и сообщений из родительской страницы в дочернее диалоговое окно осуществляется с помощью `messageChild` метода, как описано в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)надстройках Office. Если ваша надстройка работает на платформе или узле, которая не поддерживает [набор требований DialogApi 1,2](../reference/requirement-sets/dialog-api-requirement-sets.md), существует два других способа передачи сведений в диалоговое окно:</span><span class="sxs-lookup"><span data-stu-id="3b686-104">The recommended way to pass data and messages from a parent page to a child dialog box is with the `messageChild` method as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box). If your add-in is running on a platform or host that does not support the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md), there are two other ways that you can pass information to the dialog box:</span></span>

- <span data-ttu-id="3b686-105">Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="3b686-105">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="3b686-106">Храните информацию в месте, доступном как для главного, так и для диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="3b686-106">Store the information somewhere that is accessible to both the host window and dialog box.</span></span> <span data-ttu-id="3b686-107">Два окна не используют общее хранилище сеанса (свойство [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ), но *если они имеют один и тот же домен* (включая номер порта, если они есть), они совместно используют общее [Локальное хранилище](https://www.w3schools.com/html/html5_webstorage.asp).\*</span><span class="sxs-lookup"><span data-stu-id="3b686-107">The two windows do not share a common session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), but *if they have the same domain* (including port number, if any), they share a common [Local Storage](https://www.w3schools.com/html/html5_webstorage.asp).\*</span></span>


> [!NOTE]
> <span data-ttu-id="3b686-108">\* Существует ошибка, влияющая на вашу стратегию обработки маркеров.</span><span class="sxs-lookup"><span data-stu-id="3b686-108">\* There is a bug that will effect your strategy for token handling.</span></span> <span data-ttu-id="3b686-109">Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.</span><span class="sxs-lookup"><span data-stu-id="3b686-109">If the add-in is running in **Office on the web** in either the Safari or Edge browser, the dialog box and task pane do not share the same Local Storage, so it cannot be used to communicate between them.</span></span>

## <a name="use-local-storage"></a><span data-ttu-id="3b686-110">Использование локального хранилища</span><span class="sxs-lookup"><span data-stu-id="3b686-110">Use local storage</span></span>

<span data-ttu-id="3b686-111">Чтобы использовать локальное хранилище, вызовите `setItem` метод `window.localStorage` объекта на главной странице перед `displayDialogAsync` вызовом, как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="3b686-111">To use local storage, call the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="3b686-112">Код в диалоговом окне считывает элемент, когда он необходим, как в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="3b686-112">Code in the dialog box reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a><span data-ttu-id="3b686-113">Использование параметров запроса</span><span class="sxs-lookup"><span data-stu-id="3b686-113">Use query parameters</span></span>

<span data-ttu-id="3b686-114">В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.</span><span class="sxs-lookup"><span data-stu-id="3b686-114">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="3b686-115">Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span><span class="sxs-lookup"><span data-stu-id="3b686-115">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="3b686-116">Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.</span><span class="sxs-lookup"><span data-stu-id="3b686-116">Code in your dialog box can parse the URL and read the parameter value.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3b686-117">Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="3b686-117">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="3b686-118">(Этот параметр добавляется после пользовательских параметров запроса, если они есть.</span><span class="sxs-lookup"><span data-stu-id="3b686-118">(It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="3b686-119">Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать.</span><span class="sxs-lookup"><span data-stu-id="3b686-119">It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it.</span></span> <span data-ttu-id="3b686-120">Одно и то же значение добавляется в хранилище сеанса диалогового окна (свойство [Window. sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) ).</span><span class="sxs-lookup"><span data-stu-id="3b686-120">The same value is added to the dialog box's session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="3b686-121">*Ваш код не должен ни считывать это значение, ни записывать в него данные*.</span><span class="sxs-lookup"><span data-stu-id="3b686-121">Again, *your code should neither read nor write to this value*.</span></span>
