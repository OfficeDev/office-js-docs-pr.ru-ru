---
title: Добавление поддержки мобильных устройств в надстройку Outlook
description: Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: de5f1209527c853afb2d0bf2061bd3e3cfa8d3e0
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/10/2020
ms.locfileid: "43225668"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>Добавление поддержки команд надстроек для Outlook Mobile

С помощью команд надстройки в Outlook Mobile пользователи могут получить доступ к одной и той же функции (с некоторыми [ограничениями](#code-considerations)), которые уже есть в Outlook в Интернете, Windows и Mac. Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.

## <a name="updating-the-manifest"></a>Обновление манифеста

Чтобы включить команды надстроек в Outlook Mobile, необходимо сначала определить их в манифесте надстройки. В схеме [VersionOverrides](../reference/manifest/versionoverrides.md) версии 1.1 определен новый форм-фактор для мобильных устройств — [MobileFormFactor](../reference/manifest/mobileformfactor.md).

Этот элемент содержит все данные для загрузки надстройки в мобильных клиентах. Это позволяет определять совершенно другие элементы пользовательского интерфейса и файлы JavaScript для мобильной версии.

В приведенном ниже примере показана одна кнопка области задач `MobileFormFactor` в элементе.

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

Она во многом подобна элементам, которые отображаются в элементе [DesktopFormFactor](../reference/manifest/desktopformfactor.md), но имеет некоторые существенные отличия.

- Элемент [OfficeTab](../reference/manifest/officetab.md) не используется.
- У элемента [ExtensionPoint](../reference/manifest/extensionpoint.md) должен быть только один дочерний элемент. Если надстройка добавляет только одну кнопку, это должен быть дочерний элемент [Control](../reference/manifest/control.md). Если же надстройка добавляет несколько кнопок, это должен быть дочерний элемент [Group](../reference/manifest/group.md), содержащий несколько элементов `Control`.
- Для элемента `Menu` нет аналога типа `Control`.
- Элемент [Supertip](../reference/manifest/supertip.md) не используется.
- Требуются значки других размеров. Мобильные надстройки должны поддерживать как минимум значки размерами 25x25, 32x32 и 48x48 пикселей.

## <a name="code-considerations"></a>Особенности кода

При разработке надстроек для мобильных устройств возникают некоторые дополнительные особенности.

### <a name="use-rest-instead-of-exchange-web-services"></a>Использование REST вместо веб-служб Exchange

Метод [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) не поддерживается в Outlook Mobile. По мере возможности надстройки должны отдавать предпочтение данным из API Office.js. Если надстройкам требуются сведения, которые не предоставляет API Office.js, то для доступа к почтовому ящику пользователя следует использовать [интерфейсы REST API Outlook](/outlook/rest/).

В наборе требований к почтовому ящику 1,5 появилась новая версия [Office. Context. Mailbox. getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) , которая может запрашивать маркер доступа, совместимый с REST API, и новое свойство [Office. Context. Mailbox. рестурл](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) , которое можно использовать для поиска конечной точки REST API для пользователя.

### <a name="pinch-zoom"></a>Масштабирование жестами

По умолчанию пользователи могут приближать области задач с помощью жеста масштабирования. Если в вашем случае это неуместно, отключите масштабирование жестами в коде HTML.

### <a name="close-task-panes"></a>Закрытие области задач

В Outlook Mobile области задач занимают весь экран, поэтому для возврата к сообщению их необходимо закрывать. Рекомендуем использовать метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--), чтобы закрыть область задач по завершении сценария.

### <a name="compose-mode-and-appointments"></a>Режим создания и встречи

В настоящее время надстройки Outlook Mobile поддерживают активацию только при просмотре сообщений. Надстройки не активируются при создании сообщений, а также при просмотре и создании встреч. Однако интегрированные надстройки поставщика собраний по сети можно активировать в режиме организатора встреч. Дополнительные сведения об этом исключении вы найдете в статье [Создание надстройки Outlook для мобильных устройств для поставщика собраний по сети](online-meeting.md) .

### <a name="unsupported-apis"></a>Неподдерживаемые интерфейсы API

Интерфейсы API, представленные в наборе требований 1,6 или более поздних версий, не поддерживаются в Outlook Mobile. Также не поддерживаются следующие API из более ранних наборов обязательных элементов.

  - [Office.context.officeTheme](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [Office.context.mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [Office.context.mailbox.convertToEwsId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [Office.context.mailbox.item.displayReplyAllForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.displayReplyForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a>См. также

[Поддержка набора требований](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)