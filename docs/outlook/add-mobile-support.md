---
title: Добавление поддержки мобильных устройств в надстройку Outlook
description: Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 451476e9cd7eda9902ad156966558c182752547f
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484306"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Добавление поддержки команд надстроек для Outlook Mobile

Использование команд надстройки в Outlook Mobile позволяет пользователям получать доступ к той же функциональности (с некоторыми ограничениями[), что](#code-considerations) у них уже есть в Outlook в Интернете, Windows и Mac. Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.

## <a name="updating-the-manifest"></a>Обновление манифеста

Чтобы включить команды надстроек в Outlook Mobile, необходимо сначала определить их в манифесте надстройки. В схеме [VersionOverrides](/javascript/api/manifest/versionoverrides) версии 1.1 определен новый форм-фактор для мобильных устройств — [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

Этот элемент содержит все данные для загрузки надстройки в мобильных клиентах. Это позволяет определять совершенно другие элементы пользовательского интерфейса и файлы JavaScript для мобильной версии.

В приведенном ниже примере показана одна кнопка области задач в элементе `MobileFormFactor`.

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
- У элемента [ExtensionPoint](/javascript/api/manifest/extensionpoint) должен быть только один дочерний элемент. Если надстройка добавляет только одну кнопку, это должен быть дочерний элемент [Control](/javascript/api/manifest/control). Если же надстройка добавляет несколько кнопок, это должен быть дочерний элемент [Group](/javascript/api/manifest/group), содержащий несколько элементов `Control`.
- Для элемента `Menu` нет аналога типа `Control`.
- Элемент [Supertip](/javascript/api/manifest/supertip) не используется.
- Требуются значки других размеров. Мобильные надстройки должны поддерживать как минимум значки размерами 25x25, 32x32 и 48x48 пикселей.

## <a name="code-considerations"></a>Особенности кода

При разработке надстроек для мобильных устройств возникают некоторые дополнительные особенности.

### <a name="use-rest-instead-of-exchange-web-services"></a>Использование REST вместо веб-служб Exchange

Метод [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) не поддерживается в Outlook Mobile. По мере возможности надстройки должны отдавать предпочтение данным из API Office.js. Если надстройкам требуются сведения, которые не предоставляет API Office.js, то для доступа к почтовому ящику пользователя следует использовать [интерфейсы REST API Outlook](/outlook/rest/).

Набор требований к почтовым ящикам 1.5 представил новую версию [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods), которая может запрашивать маркер доступа, совместимый с API REST, и новое [свойство Office.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties), которое можно использовать для поиска конечной точки API REST для пользователя.

### <a name="pinch-zoom"></a>Масштабирование жестами

По умолчанию пользователи могут приближать области задач с помощью жеста масштабирования. Если в вашем случае это неуместно, отключите масштабирование жестами в коде HTML.

### <a name="close-task-panes"></a>Закрытие области задач

В Outlook Mobile области задач занимают весь экран, поэтому для возврата к сообщению их необходимо закрывать. Рекомендуем использовать метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)), чтобы закрыть область задач по завершении сценария.

### <a name="compose-mode-and-appointments"></a>Режим создания и встречи

В настоящее время надстройки в Outlook Mobile поддерживают активацию только при чтении сообщений. Надстройки не активируются при создании сообщений, а также при просмотре и создании встреч. Однако интегрированные надстройки поставщика собраний в Интернете можно активировать в режиме Организатор встречи. Дополнительные данные об этом исключении (включая доступные API) см. в Outlook мобильной надстройки для поставщика [онлайн-собраний](online-meeting.md#available-apis).

### <a name="unsupported-apis"></a>Неподдерживаемые интерфейсы API

API, введенные в наборе требований 1.6 или более поздней, не поддерживаются Outlook Mobile. Следующие API из предыдущих наборов требований также не поддерживаются.

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

[Наборы обязательных элементов, поддерживаемые серверами Exchange и клиентами Outlook](/javascript/api/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)