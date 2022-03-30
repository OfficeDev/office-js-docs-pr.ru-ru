---
title: Интеграция встроенных Office в настраиваемые группы управления и вкладки
description: Узнайте, как включить встроенные кнопки Office в настраиваемые группы команд и вкладки на Office ленте.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 91e64e3939ea83c6468b1f8b35ac189ad7d3d373
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496728"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Интеграция встроенных Office в настраиваемые группы управления и вкладки

Встроенные кнопки Office в настраиваемые группы управления на ленте Office с помощью разметки в манифесте надстройки. (Вы не можете вставить настраиваемые команды надстройки во встроенную Office группу.) Вы также можете вставить целые встроенные группы Office в пользовательские вкладки ленты.

> [!NOTE]
> В этой статье предполагается, что вы знакомы со статьей [Основные понятия для команд надстройки](add-in-commands.md). Пожалуйста, просмотрите его, если вы еще не сделали этого в последнее время.

> [!IMPORTANT]
>
> - Функция надстройки и разметка, описанные в этой статье, доступна только в *PowerPoint в Интернете*.
> - Разметка, описанная в этой статье, работает только на платформах, поддерживаюх набор требований **AddinCommands 1.3**. См. в более позднем [разделе Поведение на неподтверченных платформах](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Вставьте встроенную группу управления в настраиваемую вкладку

Чтобы вставить встроенную группу Office в вкладку, добавьте элемент [OfficeGroup](/javascript/api/manifest/customtab#officegroup) в качестве детского элемента родительского **элемента CustomTab**. Атрибут `id` элемента **OfficeGroup** установлен в ID встроенной группы. См [. в рублях Найти ID элементов управления и групп управления](#find-the-ids-of-controls-and-control-groups).

В следующем примере разметки группа управления Office абзаца добавляется в настраиваемую вкладку и позиционет ее, чтобы она появляться сразу после настраиваемой группы.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Вставьте встроенный контроль в настраиваемую группу

Чтобы вставить встроенный элемент Office в настраиваемую группу, добавьте элемент [OfficeControl](/javascript/api/manifest/group#officecontrol) в качестве детского элемента в элемент **родительской группы**. Атрибут `id` элемента **OfficeControl** установлен в ID встроенного элемента управления. См [. в рублях Найти ID элементов управления и групп управления](#find-the-ids-of-controls-and-control-groups).

В следующем примере разметки добавляется Office superscript в настраиваемую группу и он должен отображаться сразу после настраиваемой кнопки.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
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

Если надстройка установлена на платформе, которая не поддерживает набор требований [AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), то разметка, описанная в этой статье, игнорируется, а встроенные элементы управления и группы Office не отображаются в настраиваемой группе или вкладке. Чтобы не допустить установки надстройки на платформах, которые не поддерживают разметку, добавьте ссылку на набор требований в разделе **Требования** манифеста. Инструкции см. в Office версии и платформы, на которых может быть организована [надстройка](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Кроме того, спроектировать надстройку, чтобы иметь опыт, когда **AddinCommands 1.3** не поддерживается, как описано в [Design для](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences) альтернативных опытом. Например, если надстройка содержит инструкции, предполагаемые, что встроенные кнопки находятся в настраиваемой группе, можно создать версию, предполагаемую, что встроенные кнопки находятся только в обычных местах.
