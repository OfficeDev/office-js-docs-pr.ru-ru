---
title: Get and set internet headers
description: Как получить и установить интернет-заготки в сообщении в Outlook надстройки.
ms.date: 04/28/2020
ms.localizationpriority: medium
---

# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Получите и установите в надстройки сообщение в Outlook в интернете

## <a name="background"></a>Общие сведения

Обычное требование Outlook разработке надстройки — хранить настраиваемые свойства, связанные с надстройки на разных уровнях. В настоящее время настраиваемые свойства хранятся на уровне элемента или почтового ящика.

- Уровень элемента . Для свойств, применимых к определенному элементу, используйте [объект CustomProperties](/javascript/api/outlook/office.customproperties) . Например, храним код клиента, связанный с человеком, отправив сообщение.
- Уровень почтовых ящиков . Для свойств, применимых ко всем почтовым пунктам в почтовом ящике пользователя, используйте объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Например, храните предпочтения пользователя, чтобы показать температуру в определенной шкале.

Оба типа свойств не сохраняются после того, как элемент покидает сервер Exchange, чтобы получатели электронной почты не могли получить какие-либо свойства, установленные на элементе. Поэтому разработчики не могут получить доступ к этим настройкам или другим свойствам MIME, чтобы включить сценарии лучшего чтения.

Хотя существует способ настроить интернет-заготки с помощью запросов EWS, в некоторых случаях запрос EWS не будет работать. Например, в режиме Compose Outlook рабочем столе id `saveAsync`  элемента не синхронизирован в кэшном режиме onin.

> [!TIP]
> Дополнительные данные об использовании этих параметров см. в Outlook надстройки get and set [add-in](metadata-for-an-outlook-add-in.md).

## <a name="purpose-of-the-internet-headers-api"></a>Назначение API для интернет-заголовок

Введенные [в набор требований 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), API-API для интернет-заготчиков позволяют разработчикам:

- Печать сведений об электронной почте, которая сохраняется после Exchange всех клиентов.
- Ознакомьтесь с информацией по электронной почте, которая сохранялась после Exchange всех клиентов в сценариях чтения почты.
- Доступ ко всему заглаву MIME электронной почты.

![Схема интернет-заготов. Текст. Пользователь 1 отправляет электронную почту. Надстройка управляет настраиваемой интернет-загонами, пока пользователь создает электронную почту. Пользователь 2 получает сообщение электронной почты. Надстройка получает интернет-заготки от полученной электронной почты, а затем разбирается и использует настраиваемые заглавные.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Настройка интернет-заготов при сочинении сообщения

Попробуйте использовать [свойство item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) для управления настраиваемые интернет-заголовок, которые вы поместите в текущем сообщении в режиме Compose.

### <a name="set-get-and-remove-custom-headers-example"></a>Установите, получите и удалите пример настраиваемой заготки

В следующем примере показано, как настроить, получить и удалить настраиваемые загона.

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

## <a name="get-internet-headers-while-reading-a-message"></a>Получить заготки в Интернете при чтении сообщения

Попробуйте [вызывать item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) , чтобы получить интернет-заголовок текущего сообщения в режиме Чтения.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Получите предпочтения отправителей из текущего примера заглавных окей MIME

На примере предыдущего раздела в следующем коде показано, как получить предпочтения отправителей из заглавных записей MIME текущего адреса электронной почты.

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
> Этот пример работает для простых случаев. Для получения более сложных сведений (например, нескольких экземпляров или сложенных значений, описанных в [RFC 2822](https://tools.ietf.org/html/rfc2822)), попробуйте использовать соответствующую библиотеку для разрисовки MIME.

## <a name="recommended-practices"></a>Рекомендации

В настоящее время интернет-заготки являются конечным ресурсом в почтовом ящике пользователя. Когда квота исчерпана, вы не сможете создать в этом почтовом ящике больше интернет-заголовок, что может привести к неожиданному поведению клиентов, которые полагаются на это, чтобы функционировать.

При создании надстройки в интернете при создании надстройки применяются следующие рекомендации.

- Создайте минимальное количество необходимых загодеров.
- Заглавные имена, чтобы можно было повторно использовать и обновлять их значения позже. Таким образом, избегайте именования заглавных имен в переменной (например, на основе ввода пользователя, timestamp и т.д.).

## <a name="see-also"></a>См. также

- [Просмотр и изменение метаданных для надстройки Outlook](metadata-for-an-outlook-add-in.md)
