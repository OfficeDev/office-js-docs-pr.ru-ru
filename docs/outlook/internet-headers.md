---
title: Получение и установка заголовков в Интернете
description: Как получить и задать заголовки Интернета в сообщении в надстройке Outlook.
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8e4af70b24a96b8d00acc7ea4101acf53e2b71
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616030"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Получение и установка заголовков Интернета в сообщении в надстройке Outlook

## <a name="background"></a>Общие сведения

Распространенным требованием при разработке надстроек Outlook является хранение пользовательских свойств, связанных с надстройки, на разных уровнях. В настоящее время пользовательские свойства хранятся на уровне элемента или почтового ящика.

- Уровень элемента. Для свойств, применимых к определенному элементу, используйте объект [CustomProperties](/javascript/api/outlook/office.customproperties) . Например, сохраните код клиента, связанный с человеком, отправивший сообщение электронной почты.
- Уровень почтового ящика. Для свойств, которые применяются ко всем почтовым элементам в почтовом ящике пользователя, используйте объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Например, сохраните предпочтения пользователя, чтобы отобразить температуру в определенном масштабе.

Оба типа свойств не сохраняются после того, как элемент покидает сервер Exchange Server, поэтому получатели электронной почты не могут получить какие-либо свойства, заданные для элемента. Поэтому разработчики не могут получить доступ к этим параметрам или другим свойствам MIME для улучшения сценариев чтения.

Хотя существует способ задать заголовки Интернета с помощью запросов веб-служб Exchange (EWS), в некоторых сценариях выполнение запроса EWS не будет работать. Например, в режиме создания на рабочем столе Outlook идентификатор элемента не синхронизируется в `saveAsync` кэшированном режиме.

> [!TIP]
> Дополнительные сведения об использовании этих параметров см. в статье "Получение и установка метаданных надстройки [для надстройки Outlook"](metadata-for-an-outlook-add-in.md).

## <a name="purpose-of-the-internet-headers-api"></a>Назначение API заголовков Интернета

Представленные [в наборе обязательных элементов 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) API заголовков Интернета позволяют разработчикам:

- Пометка сведений в сообщении электронной почты, которое сохраняется после того, как оно покидает Exchange на всех клиентах.
- Чтение сведений о сообщении электронной почты, сохраненном после того, как сообщение электронной почты покинуло Exchange для всех клиентов в сценариях чтения почты.
- Доступ ко всему заголовку MIME сообщения электронной почты.

![Схема заголовков Интернета. Текст: пользователь 1 отправляет сообщение электронной почты. Надстройка управляет настраиваемыми заголовками в Интернете, пока пользователь создает электронную почту. Пользователь 2 получает сообщение электронной почты. Надстройка получает заголовки Интернета из полученной электронной почты, а затем анализирует и использует пользовательские заголовки.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Настройка заголовков Интернета при создании сообщения

Используйте свойство [item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) для управления настраиваемыми заголовками Интернета, размещаемыми в текущем сообщении в режиме создания.

### <a name="set-get-and-remove-custom-internet-headers-example"></a>Пример задания, получения и удаления настраиваемых заголовков Интернета

В следующем примере показано, как задать, получить и удалить настраиваемые заголовки Интернета.

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

[Вызовите item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)), чтобы получить заголовки Интернета для текущего сообщения в режиме чтения.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Получение параметров отправителя из текущего примера заголовков MIME

На основе примера из предыдущего раздела в следующем коде показано, как получить параметры отправителя из заголовков MIME текущего адреса электронной почты.

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
> Этот пример подходит для простых случаев. Для получения более сложных сведений (например, заголовков с несколькими экземплярами или сложенных значений, как описано в [RFC 2822](https://tools.ietf.org/html/rfc2822)), попробуйте использовать соответствующую библиотеку синтаксического анализа MIME.

## <a name="recommended-practices"></a>Рекомендации

В настоящее время заголовки Интернета являются конечным ресурсом в почтовом ящике пользователя. Если квота исчерпана, вы не сможете создать в этом почтовом ящике дополнительные заголовки Интернета, что может привести к неожиданному поведению клиентов, которые используют эту функцию.

При создании заголовков Интернета в надстройке следуйте приведенным ниже рекомендациям.

- Создайте минимальное необходимое количество заголовков. Квота заголовка зависит от общего размера заголовков, примененных к сообщению. В Exchange Online ограничение заголовка ограничивается 256 КБ, а в локальной среде Exchange ограничение определяется администратором вашей организации. Дополнительные сведения об ограничениях заголовков см. [](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits) в Exchange Online ограничениях сообщений и Exchange Server [сообщений](/exchange/mail-flow/message-size-limits).
- Имена заголовков, чтобы вы могли повторно использовать и обновлять их значения позже. Поэтому избегайте именования заголовков переменным образом (например, на основе введенных пользователем данных, метки времени и т. д.).

## <a name="see-also"></a>См. также

- [Просмотр и изменение метаданных для надстройки Outlook](metadata-for-an-outlook-add-in.md)
