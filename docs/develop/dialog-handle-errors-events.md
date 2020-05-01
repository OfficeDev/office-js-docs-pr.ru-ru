---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Описывает перехват и обработку ошибок при открытии и использовании диалогового окна "Office"
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: a35131a46dc9f5edc18df37495abe5d8c2c5ad2a
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650121"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="a89ea-103">Обработка ошибок и событий в диалоговом окне "Office"</span><span class="sxs-lookup"><span data-stu-id="a89ea-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="a89ea-104">В этой статье описывается, как выполнять перехват и обработку ошибок при открытии диалогового окна и ошибок, происходящих в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="a89ea-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="a89ea-105">В этой статье предсказано, что вы знакомы с основами использования API диалоговых окон Office, описанных в статье [Использование API диалоговых окон Office в](dialog-api-in-office-add-ins.md)надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="a89ea-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="a89ea-106">Кроме того, вы можете ознакомиться [с рекомендациями и правилами для API диалоговых окон Office](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="a89ea-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="a89ea-107">Код должен обрабатывать две категории событий:</span><span class="sxs-lookup"><span data-stu-id="a89ea-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="a89ea-108">Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="a89ea-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="a89ea-109">Ошибки и другие события в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="a89ea-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="a89ea-110">Ошибки метода displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="a89ea-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="a89ea-111">В дополнение к общим ошибкам платформы и системы, четыре ошибки относятся к вызову `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="a89ea-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="a89ea-112">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="a89ea-112">Code number</span></span>|<span data-ttu-id="a89ea-113">Значение</span><span class="sxs-lookup"><span data-stu-id="a89ea-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="a89ea-114">12004</span><span class="sxs-lookup"><span data-stu-id="a89ea-114">12004</span></span>|<span data-ttu-id="a89ea-p101">Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).</span><span class="sxs-lookup"><span data-stu-id="a89ea-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="a89ea-117">12005</span><span class="sxs-lookup"><span data-stu-id="a89ea-117">12005</span></span>|<span data-ttu-id="a89ea-118">URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP.</span><span class="sxs-lookup"><span data-stu-id="a89ea-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="a89ea-119">Необходим протокол HTTPS.</span><span class="sxs-lookup"><span data-stu-id="a89ea-119">HTTPS is required.</span></span> <span data-ttu-id="a89ea-120">(В некоторых версиях Office текст сообщения об ошибке, возвращенный с 12005, совпадает с указанным для 12004.)</span><span class="sxs-lookup"><span data-stu-id="a89ea-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="a89ea-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="a89ea-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="a89ea-p103">Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="a89ea-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="a89ea-124">12009</span><span class="sxs-lookup"><span data-stu-id="a89ea-124">12009</span></span>|<span data-ttu-id="a89ea-125">Пользователь проигнорировал диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="a89ea-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="a89ea-126">Эта ошибка может возникать в Office в Интернете, где пользователи могут отказаться от того, чтобы надстройка не могла показать диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="a89ea-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="a89ea-127">Дополнительные сведения см в разделе [Обработка блокирования всплывающих окон с помощью Office в Интернете](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a89ea-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="a89ea-128">Когда `displayDialogAsync` вызывается, объект [asyncResult](/javascript/api/office/office.asyncresult) передается в функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a89ea-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="a89ea-129">При успешном вызове открывается диалоговое окно, и `value` свойство `AsyncResult` объекта является объектом [диалогового окна](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="a89ea-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="a89ea-130">Например, в [диалоговом окне "отправить сведения" на страницу узла](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span><span class="sxs-lookup"><span data-stu-id="a89ea-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="a89ea-131">Когда вызов завершается `displayDialogAsync` с ошибкой, диалоговое окно не создается `status` , свойству `AsyncResult` объекта присваивается значение `Office.AsyncResultStatus.Failed`, и `error` свойство объекта заполняется.</span><span class="sxs-lookup"><span data-stu-id="a89ea-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="a89ea-132">Всегда следует предоставлять обратный вызов, который проверяет `status` и отвечает на сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="a89ea-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="a89ea-133">Пример, в котором сообщается о сообщении об ошибке независимо от его кода, представлен в приведенном ниже коде.</span><span class="sxs-lookup"><span data-stu-id="a89ea-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="a89ea-134">( `showNotification` Функция, не определенная в этой статье, либо отображает ошибку, либо заносит ее в журнал.</span><span class="sxs-lookup"><span data-stu-id="a89ea-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="a89ea-135">Пример реализации этой функции в надстройке приведен в статье [Пример использования API диалоговых окон надстроек Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="a89ea-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="a89ea-136">Ошибки и события в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="a89ea-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="a89ea-137">Три ошибки и события в диалоговом окне вызывают `DialogEventReceived` событие на главной странице.</span><span class="sxs-lookup"><span data-stu-id="a89ea-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="a89ea-138">Напоминание о странице ведущего приложения можно узнать в разделе [Открытие диалогового окна на странице узла](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="a89ea-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="a89ea-139">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="a89ea-139">Code number</span></span>|<span data-ttu-id="a89ea-140">Значение</span><span class="sxs-lookup"><span data-stu-id="a89ea-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="a89ea-141">12002</span><span class="sxs-lookup"><span data-stu-id="a89ea-141">12002</span></span>|<span data-ttu-id="a89ea-142">Одно из следующих:</span><span class="sxs-lookup"><span data-stu-id="a89ea-142">One of the following:</span></span><br> <span data-ttu-id="a89ea-143">– По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.</span><span class="sxs-lookup"><span data-stu-id="a89ea-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="a89ea-144">— Страница, которая была `displayDialogAsync` перезагружена, но диалоговое окно было перенаправлено на страницу, которая не может быть найдена или загружена, или она направлена на URL-адрес с недопустимым синтаксисом.</span><span class="sxs-lookup"><span data-stu-id="a89ea-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="a89ea-145">12003</span><span class="sxs-lookup"><span data-stu-id="a89ea-145">12003</span></span>|<span data-ttu-id="a89ea-p107">Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.</span><span class="sxs-lookup"><span data-stu-id="a89ea-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="a89ea-148">12006</span><span class="sxs-lookup"><span data-stu-id="a89ea-148">12006</span></span>|<span data-ttu-id="a89ea-149">Диалоговое окно было закрыто, как правило, потому что пользователь выбрал кнопку **закрытия** **X**.</span><span class="sxs-lookup"><span data-stu-id="a89ea-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="a89ea-p108">Код может назначить обработчик для события `DialogEventReceived` при вызове `displayDialogAsync`. Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="a89ea-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="a89ea-152">Ниже приведен пример обработчика для события `DialogEventReceived`, который создает особые сообщения об ошибках для каждого кода ошибки.</span><span class="sxs-lookup"><span data-stu-id="a89ea-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="a89ea-153">Надстройку с такой обработкой ошибок см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="a89ea-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
