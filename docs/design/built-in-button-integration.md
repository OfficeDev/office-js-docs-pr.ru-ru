---
title: Интеграция встроенных Office в настраиваемые группы управления и вкладки
description: Узнайте, как включить встроенные кнопки Office в настраиваемые группы команд и вкладки на Office ленте.
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: a32b2831bdbd9a260acd7f2c09b6a6c2d8ff5789807ac6e6e1e1073a0953fc18
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57082150"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Интеграция встроенных Office в настраиваемые группы управления и вкладки

Встроенные кнопки Office в настраиваемые группы управления на ленте Office с помощью разметки в манифесте надстройки. (Вы не можете вставить настраиваемые команды надстройки во встроенную Office группу.) Вы также можете вставить целые встроенные Office группы управления в настраиваемые вкладки ленты.

> [!NOTE]
> В этой статье предполагается, что вы знакомы со статьей Основные понятия для команд [надстройки](add-in-commands.md). Пожалуйста, просмотрите его, если вы еще не сделали этого в последнее время.

> [!IMPORTANT]
>
> - Функция надстройки и разметка, описанные в этой статье, доступна только в *PowerPoint в Интернете*.
> - Разметка, описанная в этой статье, работает только на платформах, поддерживаюх набор **требований AddinCommands 1.3.** См. в более позднем [разделе Поведение на неподтверченных платформах.](#behavior-on-unsupported-platforms)

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Вставьте встроенную группу управления в настраиваемую вкладку

Чтобы вставить встроенную группу управления Office в вкладку, добавьте элемент [OfficeGroup](../reference/manifest/customtab.md#officegroup) в родительский элемент в качестве детского `<CustomTab>` элемента. Атрибут `id` элемента `<OfficeGroup>` задается iD встроенной группы. См. [в рублях Find the IDs of controls and control groups.](#find-the-ids-of-controls-and-control-groups)

В следующем примере разметки группа управления Office абзаца добавляется в настраиваемую вкладку и позиционет ее, чтобы она появляться сразу после настраиваемой группы.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Вставьте встроенный контроль в настраиваемую группу

Чтобы вставить встроенный элемент Office в настраиваемую группу, добавьте элемент [OfficeControl](../reference/manifest/group.md#officecontrol) в родительский `<Group>` элемент. Атрибут `id` элемента `<OfficeControl>` задается iD встроенного элемента управления. См. [в рублях Find the IDs of controls and control groups.](#find-the-ids-of-controls-and-control-groups)

В следующем примере разметки Office управления Superscript в настраиваемую группу и он должен отображаться сразу после настраиваемой кнопки.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> Пользователи могут настроить ленту в Office приложении. Любые настройки пользователя переопределяют параметры манифеста. Например, пользователь может удалить кнопку из любой группы и удалить любую группу со вкладки.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Поиск ID-элементов групп управления и управления

ID для поддерживаемых групп управления и управления находятся в файлах в [ID Office управления](https://github.com/OfficeDev/office-control-ids). Следуйте инструкциям в файле ReadMe этого репо.

## <a name="behavior-on-unsupported-platforms"></a>Поведение на неподтверченных платформах

Если надстройка установлена на платформе, которая не поддерживает набор требований [AddinCommands 1.3,](../reference/requirement-sets/add-in-commands-requirement-sets.md)то разметка, описанная в этой статье, игнорируется, а встроенные элементы управления и группы Office не отображаются в настраиваемой группе или вкладке. Чтобы не допустить установки надстройки на платформах, которые не поддерживают разметку, добавьте ссылку на набор требований в разделе `<Requirements>` манифест. Инструкции см. в [элементе Set the Requirements in the manifest.](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) Кроме того, вы можете создать надстройку, чтобы иметь альтернативный опыт, когда **AddinCommands 1.3** не поддерживается, как описано в описании Использования проверок времени запуска в коде [JavaScript.](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) Например, если надстройка содержит инструкции, предполагаемые, что встроенные кнопки находятся в настраиваемой группе, вы можете иметь альтернативную версию, предполагаемую, что встроенные кнопки находятся только в обычных местах.
