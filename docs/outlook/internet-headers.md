---
title: Получение и Настройка заголовков Интернета
description: Получение и Настройка заголовков Интернета для сообщения в надстройке Outlook.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: 488a4414580296da59eef3eb703e1c8da7e7d7c2
ms.sourcegitcommit: 231e23d72e04e0536480d6b16df95113f1eff738
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43238220"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Получение и Настройка заголовков Интернета для сообщения в надстройке Outlook

## <a name="background"></a>Общие сведения

Общее требование при разработке надстроек Outlook — Хранение настраиваемых свойств, связанных с надстройкой, на различных уровнях. В настоящее время настраиваемые свойства хранятся на уровне элемента или почтового ящика.

- Уровень элемента — для свойств, которые применяются к определенному элементу, используйте объект [CustomProperties](/javascript/api/outlook/office.customproperties) . Например, сохраните код клиента, связанный с пользователем, который отправил сообщение.
- Уровень почтового ящика — для свойств, которые применяются ко всем почтовым элементам в почтовом ящике пользователя, используйте объект [roamingSettings](/javascript/api/outlook/office.roamingsettings) . Например, сохраните параметр пользователя, чтобы показать температуру в определенном масштабе.

Оба типа свойств не сохраняются после того, как элемент покидает сервер Exchange, поэтому получатели электронной почты не смогут получить никакие свойства, установленные для элемента. Таким образом, разработчики не могут получить доступ к этим параметрам или другим свойствам MIME, чтобы обеспечить более высокую удобочитаемость сценариев.

Несмотря на то, что вы можете задать заголовки Интернета с помощью запросов EWS, в некоторых случаях выполнение запроса EWS не будет работать. Например, в режиме создания на настольном компьютере Outlook идентификатор элемента не синхронизируется `saveAsync` в режиме кэширования.

> [!TIP]
> Чтобы узнать больше об использовании этих параметров, ознакомьтесь со [статьей получение и установка метаданных надстройки Outlook](metadata-for-an-outlook-add-in.md) .

## <a name="purpose-of-the-internet-headers-api"></a>Назначение API заголовков Интернета

Появилось в наборе требований 1,8, API заголовков Интернета позволяют разработчикам выполнять следующие задачи:

- Сведения о штампе в сообщении электронной почты, которое сохраняется после того, как оно покидает Exchange на всех клиентах.
- Прочитайте сведения о сообщении электронной почты, которое сохранилось после левого почтового обмена сообщениями для всех клиентов в сценариях чтения почты.
- Доступ ко всему заголовку MIME сообщения электронной почты.

![Схема заголовков Интернета. Text: пользователь 1 отправляет электронную почту. Надстройка управляет пользовательскими заголовками Интернета, когда пользователь создает электронную почту. Пользователь 2 получает сообщение электронной почты. Надстройка получает заголовки Интернета из полученного электронного письма, а затем анализирует и использует настраиваемые заголовки. ](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Настройка заголовков Интернета при создании сообщения

Попробуйте использовать свойство [Item. internetheaders:](/javascript/api/outlook/office.messagecompose#internetheaders) для управления пользовательскими заголовками Интернета, которые вы поместите в текущем сообщении в режиме создания.

### <a name="set-get-and-remove-custom-headers-example"></a>Пример задания, получения и удаления настраиваемых заголовков

В приведенном ниже примере показано, как задать, получить и удалить настраиваемые заголовки.

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

## <a name="get-internet-headers-while-reading-a-message"></a>Получение заголовков Интернета при чтении сообщения

Попробуйте вызвать [Item. жеталлинтернесеадерсасинк](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) для получения Интернет-заголовков в текущем сообщении в режиме чтения.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Пример получения параметров отправителя из текущего MIME заголовков

В примере, приведенном в предыдущем разделе, показано, как получить предпочтения отправителя из заголовков MIME текущего сообщения электронной почты.

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
> Этот пример работает в простых случаях. Для получения более сложной информации (например, многоэкземплярных заголовков или значений со сгибом, описанных в [RFC 2822](https://tools.ietf.org/html/rfc2822)) попробуйте использовать соответствующую библиотеку для синтаксического анализа MIME.

## <a name="see-also"></a>См. также

- [Просмотр и изменение метаданных для надстройки Outlook](metadata-for-an-outlook-add-in.md)
