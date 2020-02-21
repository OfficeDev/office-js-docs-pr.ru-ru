---
title: Получение и Настройка заголовков Интернета
description: Получение и Настройка заголовков Интернета для сообщения в надстройке Outlook.
ms.date: 11/04/2019
localization_priority: Normal
ms.openlocfilehash: d7f38b7564683ce51ed0bd840480b4a8b2040bf6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166754"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="f94f4-103">Получение и Настройка заголовков Интернета для сообщения в надстройке Outlook</span><span class="sxs-lookup"><span data-stu-id="f94f4-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="f94f4-104">Общие сведения</span><span class="sxs-lookup"><span data-stu-id="f94f4-104">Background</span></span>

<span data-ttu-id="f94f4-105">Общее требование при разработке надстроек Outlook — Хранение настраиваемых свойств, связанных с надстройкой, на различных уровнях.</span><span class="sxs-lookup"><span data-stu-id="f94f4-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="f94f4-106">В настоящее время настраиваемые свойства хранятся на уровне элемента или почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="f94f4-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="f94f4-107">Уровень элемента — для свойств, которые применяются к определенному элементу, используйте объект [CustomProperties](/javascript/api/outlook/office.customproperties) .</span><span class="sxs-lookup"><span data-stu-id="f94f4-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="f94f4-108">Например, сохраните код клиента, связанный с пользователем, который отправил сообщение.</span><span class="sxs-lookup"><span data-stu-id="f94f4-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="f94f4-109">Уровень почтового ящика — для свойств, которые применяются ко всем почтовым элементам в почтовом ящике пользователя, используйте объект [roamingSettings](/javascript/api/outlook/office.roamingsettings) .</span><span class="sxs-lookup"><span data-stu-id="f94f4-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="f94f4-110">Например, сохраните параметр пользователя, чтобы показать температуру в определенном масштабе.</span><span class="sxs-lookup"><span data-stu-id="f94f4-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="f94f4-111">Оба типа свойств не сохраняются после того, как элемент покидает сервер Exchange, поэтому получатели электронной почты не смогут получить никакие свойства, установленные для элемента.</span><span class="sxs-lookup"><span data-stu-id="f94f4-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="f94f4-112">Таким образом, разработчики не могут получить доступ к этим параметрам или другим свойствам MIME, чтобы обеспечить более высокую удобочитаемость сценариев.</span><span class="sxs-lookup"><span data-stu-id="f94f4-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="f94f4-113">Несмотря на то, что вы можете задать заголовки Интернета с помощью запросов EWS, в некоторых случаях выполнение запроса EWS не будет работать.</span><span class="sxs-lookup"><span data-stu-id="f94f4-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="f94f4-114">Например, в режиме создания на настольном компьютере Outlook идентификатор элемента не синхронизируется `saveAsync` в режиме кэширования.</span><span class="sxs-lookup"><span data-stu-id="f94f4-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="f94f4-115">Чтобы узнать больше об использовании этих параметров, ознакомьтесь со [статьей получение и установка метаданных надстройки Outlook](metadata-for-an-outlook-add-in.md) .</span><span class="sxs-lookup"><span data-stu-id="f94f4-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="f94f4-116">Назначение API заголовков Интернета</span><span class="sxs-lookup"><span data-stu-id="f94f4-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="f94f4-117">Появилось в наборе требований 1,8, API заголовков Интернета позволяют разработчикам выполнять следующие задачи:</span><span class="sxs-lookup"><span data-stu-id="f94f4-117">Introduced in requirement set 1.8, the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="f94f4-118">Сведения о штампе в сообщении электронной почты, которое сохраняется после того, как оно покидает Exchange на всех клиентах.</span><span class="sxs-lookup"><span data-stu-id="f94f4-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="f94f4-119">Прочитайте сведения о сообщении электронной почты, которое сохранилось после левого почтового обмена сообщениями для всех клиентов в сценариях чтения почты.</span><span class="sxs-lookup"><span data-stu-id="f94f4-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="f94f4-120">Доступ ко всему заголовку MIME сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="f94f4-120">Access the entire MIME header of the email.</span></span>

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="f94f4-121">Настройка заголовков Интернета при создании сообщения</span><span class="sxs-lookup"><span data-stu-id="f94f4-121">Set internet headers while composing a message</span></span>

<span data-ttu-id="f94f4-122">Попробуйте использовать свойство [Item. internetheaders:](/javascript/api/outlook/office.messagecompose#internetheaders) для управления пользовательскими заголовками Интернета, которые вы поместите в текущем сообщении в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="f94f4-122">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="f94f4-123">Пример задания, получения и удаления настраиваемых заголовков</span><span class="sxs-lookup"><span data-stu-id="f94f4-123">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="f94f4-124">В приведенном ниже примере показано, как задать, получить и удалить настраиваемые заголовки.</span><span class="sxs-lookup"><span data-stu-id="f94f4-124">The following example shows how to set, get, and remove custom headers.</span></span>

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="f94f4-125">Получение заголовков Интернета при чтении сообщения</span><span class="sxs-lookup"><span data-stu-id="f94f4-125">Get internet headers while reading a message</span></span>

<span data-ttu-id="f94f4-126">Попробуйте вызвать [Item. жеталлинтернесеадерсасинк](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) для получения Интернет-заголовков в текущем сообщении в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f94f4-126">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="f94f4-127">Пример получения параметров отправителя из текущего MIME заголовков</span><span class="sxs-lookup"><span data-stu-id="f94f4-127">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="f94f4-128">В примере, приведенном в предыдущем разделе, показано, как получить предпочтения отправителя из заголовков MIME текущего сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="f94f4-128">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> <span data-ttu-id="f94f4-129">Этот пример работает в простых случаях.</span><span class="sxs-lookup"><span data-stu-id="f94f4-129">This sample works for simple cases.</span></span> <span data-ttu-id="f94f4-130">Для получения более сложной информации (например, многоэкземплярных заголовков или значений со сгибом, описанных в [RFC 2822](https://tools.ietf.org/html/rfc2822)) попробуйте использовать соответствующую библиотеку для синтаксического анализа MIME.</span><span class="sxs-lookup"><span data-stu-id="f94f4-130">For more complex information retrieval (e.g., multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="see-also"></a><span data-ttu-id="f94f4-131">См. также</span><span class="sxs-lookup"><span data-stu-id="f94f4-131">See also</span></span>

- [<span data-ttu-id="f94f4-132">Просмотр и изменение метаданных для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="f94f4-132">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
