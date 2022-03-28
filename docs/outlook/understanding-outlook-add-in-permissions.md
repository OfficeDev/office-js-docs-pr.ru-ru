---
title: Общие сведения о разрешениях для надстроек Outlook
description: Надстройки Outlook указывают требуемый уровень разрешений в своем манифесте, который включает Restricted, ReadItem, ReadWriteItem, or ReadWriteMailbox.
ms.date: 02/19/2020
ms.localizationpriority: medium
ms.openlocfilehash: 6350e0d3aed499d831c13e440945fda1f60742ca
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484184"
---
# <a name="understanding-outlook-add-in-permissions"></a>Общие сведения о разрешениях для надстроек Outlook

Необходимый уровень разрешений для надстроек Outlook указывается в манифесте. Доступные уровни: **Restricted**, **ReadItem**, **ReadWriteItem** и **ReadWriteMailbox**. Эти уровни являются накопительными: **Restricted** — самый низкий уровень, каждый более высокий уровень включает разрешения более низких уровней. **ReadWriteMailbox** включает все поддерживаемые разрешения.

Вы можете просмотреть разрешения, которые запрашивает почтовая надстройка, перед ее установкой из [AppSource](https://appsource.microsoft.com). Вы также можете просмотреть требуемые разрешения установленных надстроек в Центре администрирования Exchange.

## <a name="restricted-permission"></a>Разрешение Restricted


  **Restricted** — самый простой уровень разрешений. Укажите **Restricted** в элементе [Permissions](/javascript/api/manifest/permissions) манифеста, чтобы запросить это разрешение. Outlook назначает это разрешение почтовой надстройке по умолчанию, если надстройка не запрашивает особого разрешения в манифесте.

### <a name="can-do"></a>Разрешено

- [Получать только определенные сущности](match-strings-in-an-item-as-well-known-entities.md) (номер телефона, адрес, URL-адрес) из темы или текста элемента.

- Указывать [правило активации ItemIs](activation-rules.md#itemis-rule), требующее, чтобы текущий элемент в форме чтения или создания принадлежал определенному типу, или правило [ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md), соответствующее малому поднабору поддерживаемых известных сущностей (номер телефона, адрес, URL-адрес) в выбранном элементе.

- Получать доступ к свойствам и методам, которые **не** относятся к определенной информации о пользователе или элементе (список элементов, которые относятся к такой информации, см. в следующем разделе).

### <a name="cant-do"></a>Не разрешено

- Используйте правило [ItemHasKnownEntity для](/javascript/api/manifest/rule#itemhasknownentity-rule) объекта контактов, электронной почты, предложения собрания или предложения задач.

- Использовать правило [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) или [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule).

- Получать доступ к элементам в приведенном ниже списке, которые относятся к информации о пользователе или элементе. При попытке получить доступ к элементам в этом списке будут возвращены значение **null** и сообщение о том, что требуются повышенные привилегии.

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) и все дочерние элементы
  - [Location](/javascript/api/outlook/office.location) и все дочерние элементы
  - [Recipients](/javascript/api/outlook/office.recipients) и все дочерние элементы
  - [Subject](/javascript/api/outlook/office.subject) и все дочерние элементы
  - [Time](/javascript/api/outlook/office.time) и все дочерние элементы

## <a name="readitem-permission"></a>Разрешение ReadItem

**ReadItem** — следующий уровень в модели разрешений. Укажите **ReadItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение.

### <a name="can-do"></a>Разрешено

- [Считывать все свойства](item-data.md) текущего элемента в чтении или [Создавать форму](get-and-set-item-data-in-a-compose-form.md), например [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) в форме чтения и [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1)) в форме создания.

- [Получать маркер обратного вызова для получения вложений](get-attachments-of-an-outlook-item.md) или всего элемента с помощью веб-служб Exchange или [REST API Outlook](use-rest-api.md).

- [Записывать пользовательские свойства](/javascript/api/outlook/office.customproperties), установленные надстройкой для соответствующего элемента.

- [Получать все существующие известные сущности](match-strings-in-an-item-as-well-known-entities.md) (а не только группу) из темы или текста элемента.

- Используйте все [известные](activation-rules.md#itemhasknownentity-rule) сущности в [правилах ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) или регулярных выражениях в [правилах ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule).[](activation-rules.md#itemhasregularexpressionmatch-rule) В следующем примере следует схема v1.1. В нем показано правило, которое активирует надстройки, если одна или несколько известных сущностями находятся в субъекте или теле выбранного сообщения.

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a>Не разрешено

- Использовать токен, предоставляемый методом **mailbox.getCallbackTokenAsync**, для следующего:
  - обновление или удаление текущего элемента с помощью REST API для Outlook и получение доступа к другим элементам в почтовом ящике пользователя;
  - получение текущего элемента события календаря с помощью REST API для Outlook.

- Используйте любой из следующих API.
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## <a name="readwriteitem-permission"></a>Разрешение ReadWriteItem

Укажите элемент **ReadWriteItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение. Почтовые надстройки, активированные в формах создания, которые используют методы записи (**Message.to.addAsync** или **Message.to.setAsync**), должны использовать по крайней мере этот уровень разрешений.

### <a name="can-do"></a>Разрешено

- [Считывать и записывать все свойства на уровне элемента](item-data.md) для элемента, который просматривается или создается в Outlook.

- [Добавлять или удалять вложения](add-and-remove-attachments-to-an-item-in-a-compose-form.md) для такого элемента.

- Используйте все другие члены Office JavaScript API, применимые к почтовым надстройки, за исключением **Mailbox.makeEWSRequestAsync**.

### <a name="cant-do"></a>Не разрешено

- Использовать токен, предоставляемый методом **mailbox.getCallbackTokenAsync**, для следующего:
  - обновление или удаление текущего элемента с помощью REST API для Outlook и получение доступа к другим элементам в почтовом ящике пользователя;
  - получение текущего элемента события календаря с помощью REST API для Outlook.

- Использовать **mailbox.makeEWSRequestAsync**.

## <a name="readwritemailbox-permission"></a>Разрешение ReadWriteMailbox

**ReadWriteMailbox** — самый высокий уровень разрешений. Укажите **ReadWriteMailbox** в элементе **Permissions** манифеста, чтобы запросить это разрешение.

В дополнение к тому, что поддерживает разрешение **ReadWriteItem**, токен, предоставляемый элементом **mailbox.getCallbackTokenAsync**, позволяет использовать операции веб-служб Exchange или REST API Outlook для выполнения следующих действий:

- Чтение и запись всех свойств любого элемента в почтовом ящике пользователя.
- Создание, чтение и запись в любую папку или элемент в таком почтовом ящике.
- Отправка элемента из такого почтового ящика

Через **mailbox.makeEWSRequestAsync** можно получить доступ к следующим операциям EWS.

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

Попытка использования неподдерживаемой операции приведет к возврату ошибки.

## <a name="see-also"></a>См. также

- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](../concepts/privacy-and-security.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
