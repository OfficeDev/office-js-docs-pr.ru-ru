---
title: Обработка ошибок и событий в диалоговом окне "Office"
description: Описывает, как улавливать и обрабатывать ошибки при открытии Office диалоговом окне
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349898"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="7ccff-103">Обработка ошибок и событий в диалоговом окне "Office"</span><span class="sxs-lookup"><span data-stu-id="7ccff-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="7ccff-104">В этой статье описывается, как улавливать и обрабатывать ошибки при открытии диалоговое окно и ошибки, которые происходят в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="7ccff-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="7ccff-105">В этой статье предполагается, что вы знакомы с основами использования API диалогового Office, как описано в Статье [Использование API](dialog-api-in-office-add-ins.md)диалогов Office в Office надстройки .</span><span class="sxs-lookup"><span data-stu-id="7ccff-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="7ccff-106">См. также лучшие практики и правила для API Office [диалоговом ок.](dialog-best-practices.md)</span><span class="sxs-lookup"><span data-stu-id="7ccff-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="7ccff-107">Код должен обрабатывать две категории событий:</span><span class="sxs-lookup"><span data-stu-id="7ccff-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="7ccff-108">Ошибки, возвращаемые при вызове метода `displayDialogAsync`, так как не удается создать диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7ccff-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="7ccff-109">Ошибки и другие события в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="7ccff-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="7ccff-110">Ошибки метода displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="7ccff-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="7ccff-111">Помимо общих ошибок платформы и системы, четыре ошибки являются специфическими для вызова `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="7ccff-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="7ccff-112">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="7ccff-112">Code number</span></span>|<span data-ttu-id="7ccff-113">Значение</span><span class="sxs-lookup"><span data-stu-id="7ccff-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="7ccff-114">12004</span><span class="sxs-lookup"><span data-stu-id="7ccff-114">12004</span></span>|<span data-ttu-id="7ccff-p101">Домен URL-адреса, передаваемого в метод `displayDialogAsync`, не является доверенным. Домен должен быть таким же, как и для главной страницы (а также протокол и номер порта).</span><span class="sxs-lookup"><span data-stu-id="7ccff-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="7ccff-117">12005</span><span class="sxs-lookup"><span data-stu-id="7ccff-117">12005</span></span>|<span data-ttu-id="7ccff-118">URL-адрес, передаваемый в метод `displayDialogAsync`, использует протокол HTTP.</span><span class="sxs-lookup"><span data-stu-id="7ccff-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="7ccff-119">Необходим протокол HTTPS.</span><span class="sxs-lookup"><span data-stu-id="7ccff-119">HTTPS is required.</span></span> <span data-ttu-id="7ccff-120">(В некоторых версиях Office текст сообщения об ошибке, возвращенный с 12005, является тем же, что и для 12004.)</span><span class="sxs-lookup"><span data-stu-id="7ccff-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="7ccff-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="7ccff-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="7ccff-p103">Диалоговое окно уже открыто из этого главного окна. Для главного окна, например области задач, невозможно открыть сразу несколько диалоговых окон.</span><span class="sxs-lookup"><span data-stu-id="7ccff-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="7ccff-124">12009</span><span class="sxs-lookup"><span data-stu-id="7ccff-124">12009</span></span>|<span data-ttu-id="7ccff-125">Пользователь проигнорировал диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7ccff-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="7ccff-126">Эта ошибка может возникнуть в Office в Интернете, когда пользователи могут не разрешать надстройке представлять диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="7ccff-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="7ccff-127">Дополнительные сведения см. в ссылке [Обработка всплывающих](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)блокаторов с помощью Office в Интернете .</span><span class="sxs-lookup"><span data-stu-id="7ccff-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="7ccff-128">При `displayDialogAsync` вызове он передает объект [AsyncResult](/javascript/api/office/office.asyncresult) функцию вызова.</span><span class="sxs-lookup"><span data-stu-id="7ccff-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="7ccff-129">При успешном вызове открывается диалоговое окно, а свойством объекта является `value` `AsyncResult` объект [Dialog.](/javascript/api/office/office.dialog)</span><span class="sxs-lookup"><span data-stu-id="7ccff-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="7ccff-130">В этом примере см. статью [Отправка сведений из диалогового окна на хост-страницу.](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)</span><span class="sxs-lookup"><span data-stu-id="7ccff-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="7ccff-131">При сбой вызова диалоговое окно не создается, задается свойство объекта `displayDialogAsync` `status` и `AsyncResult` `Office.AsyncResultStatus.Failed` `error` заполняется свойство объекта.</span><span class="sxs-lookup"><span data-stu-id="7ccff-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="7ccff-132">Всегда необходимо предоставить вызов, который проверяет ошибку и отвечает на `status` нее.</span><span class="sxs-lookup"><span data-stu-id="7ccff-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="7ccff-133">Пример сообщения об ошибке независимо от номера кода см. в следующем коде.</span><span class="sxs-lookup"><span data-stu-id="7ccff-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="7ccff-134">`showNotification`(Функция, не заданная в этой статье, отображает или регистрит ошибку.</span><span class="sxs-lookup"><span data-stu-id="7ccff-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="7ccff-135">Пример реализации этой функции в надстройки см. в Office примере [API диалогов надстройки.)](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)</span><span class="sxs-lookup"><span data-stu-id="7ccff-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="7ccff-136">Ошибки и события в диалоговом окне</span><span class="sxs-lookup"><span data-stu-id="7ccff-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="7ccff-137">Три ошибки и события в диалоговом окне поднимут `DialogEventReceived` событие на хост-странице.</span><span class="sxs-lookup"><span data-stu-id="7ccff-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="7ccff-138">Напоминая о том, что такое хост-страница, см. в странице Откройте диалоговое [окно с хост-страницы.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)</span><span class="sxs-lookup"><span data-stu-id="7ccff-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="7ccff-139">Цифровой код</span><span class="sxs-lookup"><span data-stu-id="7ccff-139">Code number</span></span>|<span data-ttu-id="7ccff-140">Значение</span><span class="sxs-lookup"><span data-stu-id="7ccff-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="7ccff-141">12002</span><span class="sxs-lookup"><span data-stu-id="7ccff-141">12002</span></span>|<span data-ttu-id="7ccff-142">Одно из следующих:</span><span class="sxs-lookup"><span data-stu-id="7ccff-142">One of the following:</span></span><br> <span data-ttu-id="7ccff-143">– По URL-адресу, переданному в `displayDialogAsync`, не существует страницы.</span><span class="sxs-lookup"><span data-stu-id="7ccff-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="7ccff-144">- Страница, которая была передана для загрузки, но диалоговое окно было перенаправлено на страницу, которую она не может найти или загрузить, или она была направлена на URL-адрес с недействительным `displayDialogAsync` синтаксис.</span><span class="sxs-lookup"><span data-stu-id="7ccff-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="7ccff-145">12003</span><span class="sxs-lookup"><span data-stu-id="7ccff-145">12003</span></span>|<span data-ttu-id="7ccff-p107">Выполнена попытка открыть из диалогового окна страницу, для URL-адреса которой используется протокол HTTP. Необходим протокол HTTPS.</span><span class="sxs-lookup"><span data-stu-id="7ccff-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="7ccff-148">12006</span><span class="sxs-lookup"><span data-stu-id="7ccff-148">12006</span></span>|<span data-ttu-id="7ccff-149">Диалоговое окно было закрыто, как правило, из-за того, что пользователь выбрал кнопку **Закрыть** **X**.</span><span class="sxs-lookup"><span data-stu-id="7ccff-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="7ccff-p108">Код может назначить обработчик для события `DialogEventReceived` при вызове `displayDialogAsync`. Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="7ccff-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="7ccff-152">Пример обработки события, создав настраиваемые сообщения об ошибке для каждого кода ошибки, см. `DialogEventReceived` в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="7ccff-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.</span></span>

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

<span data-ttu-id="7ccff-153">Надстройку с такой обработкой ошибок см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="7ccff-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
