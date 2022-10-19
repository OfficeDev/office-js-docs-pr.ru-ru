---
title: Добавление поддержки мобильных устройств в надстройку Outlook
description: Узнайте, как добавить поддержку Outlook Mobile, в том числе обновить манифест надстройки и при необходимости изменить код для мобильных сценариев.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c84b4aeb04cd2c8b3c2f0a7afa9fd1631c22afc5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607548"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Добавление поддержки команд надстроек для Outlook Mobile

Использование команд надстроек в Outlook Mobile позволяет пользователям получать доступ к тем же функциям (с некоторыми ограничениями[), которые](#code-considerations) у них уже есть в Outlook в Интернете, Windows и Mac. Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.

## <a name="updating-the-manifest"></a>Обновление манифеста

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

В следующем примере показана одна кнопка области задач в элементе `MobileFormFactor` .

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

Она во многом подобна элементам, которые отображаются в элементе [DesktopFormFactor](/javascript/api/manifest/desktopformfactor), но имеет некоторые существенные отличия.

- Элемент [OfficeTab](/javascript/api/manifest/officetab) не используется.
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/javascript/api/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/javascript/api/manifest/group) element that contains multiple `Control` elements.
- Для элемента `Menu` нет аналога типа `Control`.
- Элемент [Supertip](/javascript/api/manifest/supertip) не используется.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## <a name="code-considerations"></a>Особенности кода

При разработке надстроек для мобильных устройств возникают некоторые дополнительные особенности.

### <a name="use-rest-instead-of-exchange-web-services"></a>Использование REST вместо веб-служб Exchange

The [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

В наборе обязательных элементов почтового ящика 1.5 представлена новая версия [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) , которая может запрашивать маркер доступа, совместимый с REST API, и новое свойство [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) , которое можно использовать для поиска конечной точки REST API для пользователя.

### <a name="pinch-zoom"></a>Масштабирование жестами

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### <a name="close-task-panes"></a>Закрытие области задач

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### <a name="compose-mode-and-appointments"></a>Режим создания и встречи

В настоящее время надстройки в Outlook Mobile поддерживают активацию только при чтении сообщений. Надстройки не активируются при создании сообщений, а также при просмотре и создании встреч. Однако существует два исключения:

1. Встроенные надстройки поставщика собраний по сети можно активировать в режиме организатора встреч. Дополнительные сведения об этом исключении (включая доступные API) см. в статье "Создание мобильной надстройки [Outlook для поставщика онлайн-собраний"](online-meeting.md#available-apis).
1. Надстройки, которые регистрируются примечаниями к встречам и другими сведениями об управлении отношениями с клиентами (CRM) или службах создания заметок, можно активировать в режиме участника встречи. Дополнительные сведения об этом исключении (включая доступные API) см. в заметках о встрече журнала для внешнего приложения в [мобильных](mobile-log-appointments.md#available-apis) надстройки Outlook.

### <a name="unsupported-apis"></a>Неподдерживаемые интерфейсы API

API, представленные в наборе обязательных элементов 1.6 или более поздней версии, не поддерживаются Outlook Mobile. Следующие API из предыдущих наборов обязательных элементов также не поддерживаются.

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>См. также

[Наборы обязательных элементов, поддерживаемые серверами Exchange и клиентами Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)