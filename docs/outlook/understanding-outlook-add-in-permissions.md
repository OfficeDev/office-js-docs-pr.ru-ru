---
title: Общие сведения о разрешениях для надстроек Outlook
description: Надстройки Outlook указывают требуемый уровень разрешений в своем манифесте, который включает Restricted, ReadItem, ReadWriteItem, or ReadWriteMailbox.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 60b65416585b5215ed565a3689c1e7f398e001a5
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325328"
---
# <a name="understanding-outlook-add-in-permissions"></a>Общие сведения о разрешениях для надстроек Outlook

Необходимый уровень разрешений для надстроек Outlook указывается в манифесте. Доступные уровни: **Restricted**, **ReadItem**, **ReadWriteItem** и **ReadWriteMailbox**. Эти уровни являются накопительными: **Restricted** — самый низкий уровень, каждый более высокий уровень включает разрешения более низких уровней. **ReadWriteMailbox** включает все поддерживаемые разрешения.

Вы можете просмотреть разрешения, которые запрашивает почтовая надстройка, перед ее установкой из [AppSource](https://appsource.microsoft.com). Вы также можете просмотреть требуемые разрешения установленных надстроек в Центре администрирования Exchange.

## <a name="restricted-permission"></a>Разрешение Restricted


  **Restricted** — самый простой уровень разрешений. Укажите **Restricted** в элементе [Permissions](../reference/manifest/permissions.md) манифеста, чтобы запросить это разрешение. Outlook назначает это разрешение почтовой надстройке по умолчанию, если надстройка не запрашивает особого разрешения в манифесте.

### <a name="can-do"></a>Разрешено

- [Получать только определенные сущности](match-strings-in-an-item-as-well-known-entities.md) (номер телефона, адрес, URL-адрес) из темы или текста элемента.

- Указывать [правило активации ItemIs](activation-rules.md#itemis-rule), требующее, чтобы текущий элемент в форме чтения или создания принадлежал определенному типу, или правило [ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md), соответствующее малому поднабору поддерживаемых известных сущностей (номер телефона, адрес, URL-адрес) в выбранном элементе.

- Получать доступ к свойствам и методам, которые **не** относятся к определенной информации о пользователе или элементе (список элементов, которые относятся к такой информации, см. в следующем разделе).

### <a name="cant-do"></a>Не разрешено

- Используйте правило [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) для контакта, адрес электронной почты, предложение о собрании или сущность предложения по задаче.

- Использовать правило [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) или [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule).

- Получать доступ к элементам в приведенном ниже списке, которые относятся к информации о пользователе или элементе. При попытке получить доступ к элементам в этом списке будут возвращены значение **null** и сообщение о том, что требуются повышенные привилегии.

    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.userProfile](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - [Body](/javascript/api/outlook/office.body) и все дочерние элементы
    - [Location](/javascript/api/outlook/office.location) и все дочерние элементы
    - [Recipients](/javascript/api/outlook/office.recipients) и все дочерние элементы
    - [Subject](/javascript/api/outlook/office.subject) и все дочерние элементы
    - [Time](/javascript/api/outlook/office.time) и все дочерние элементы

## <a name="readitem-permission"></a>Разрешение ReadItem

**ReadItem** — следующий уровень в модели разрешений. Укажите **ReadItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение.

### <a name="can-do"></a>Разрешено

- [Считывать все свойства](item-data.md) текущего элемента в чтении или [Создавать форму](get-and-set-item-data-in-a-compose-form.md), например [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) в форме чтения и [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) в форме создания.

- [Получать маркер обратного вызова для получения вложений](get-attachments-of-an-outlook-item.md) или всего элемента с помощью веб-служб Exchange или [REST API Outlook](use-rest-api.md).

- [Записывать пользовательские свойства](/javascript/api/outlook/office.CustomProperties), установленные надстройкой для соответствующего элемента.

- [Получать все существующие известные сущности](match-strings-in-an-item-as-well-known-entities.md) (а не только группу) из темы или текста элемента.

- Использовать все [известные сущности](activation-rules.md#itemhasknownentity-rule) в правилах [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) или [регулярные выражения](activation-rules.md#itemhasregularexpressionmatch-rule) в правилах [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule). Следующий пример, использующий схему версии 1.1, активирует надстройку, если обнаруживается одна или несколько известных сущностей в теме или теле выбранного сообщения:

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

- Использовать один из следующих API:
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.bcc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.bcc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [item.body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [item.cc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.cc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.end.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.start.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [item.to.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.to.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a>Разрешение ReadWriteItem

Укажите элемент **ReadWriteItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение. Почтовые надстройки, активированные в формах создания, которые используют методы записи (**Message.to.addAsync** или **Message.to.setAsync**), должны использовать по крайней мере этот уровень разрешений.

### <a name="can-do"></a>Разрешено

- [Считывать и записывать все свойства на уровне элемента](item-data.md) для элемента, который просматривается или создается в Outlook.

- [Добавлять или удалять вложения](add-and-remove-attachments-to-an-item-in-a-compose-form.md) для такого элемента.

- Используйте все остальные элементы API JavaScript для Office, которые относятся к почтовым надстройкам, за исключением **Mailbox. makeEWSRequestAsync**.

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

С помощью **mailbox.makeEWSRequestAsync** вы можете использовать следующие операции EWS:

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

- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](../develop/privacy-and-security.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
