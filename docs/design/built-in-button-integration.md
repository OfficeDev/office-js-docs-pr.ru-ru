---
title: Интеграция встроенных кнопок Office в пользовательские группы управления и вкладки
description: Узнайте, как включить встроенные кнопки Office в пользовательские группы команд и вкладки на ленте Office.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc706fcd0b049647847a73f7c40144dba9df0e2
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659789"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>Интеграция встроенных кнопок Office в пользовательские группы управления и вкладки

Встроенные кнопки Office можно вставлять в пользовательские группы элементов управления на ленте Office с помощью разметки в манифесте надстройки. (Вы не можете вставить пользовательские команды надстройки во встроенную группу Office.) Вы также можете вставить все встроенные группы элементов управления Office на пользовательские вкладки ленты.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с основными понятиями статьи о командах [надстроек](add-in-commands.md). Ознакомьтесь с ним, если вы еще этого не сделали.

> [!IMPORTANT]
>
> - Функция надстройки и разметка, описанные в этой статье, доступны только *в PowerPoint в Интернете*.
> - Разметка, описанная в этой статье, работает только на платформах, поддерживающих набор обязательных элементов **AddinCommands 1.3**. Дополнительные сведения см. [в разделе "Поведение на неподдерживаемых платформах"](#behavior-on-unsupported-platforms).

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>Вставка встроенной группы элементов управления на настраиваемую вкладку

Чтобы вставить встроенную группу элементов управления Office на вкладку, добавьте элемент [OfficeGroup](/javascript/api/manifest/customtab#officegroup) в родительский элемент в качестве дочернего элемента **\<CustomTab\>** . Атрибуту `id` элемента **\<OfficeGroup\>** задается идентификатор встроенной группы. См [. раздел "Поиск идентификаторов элементов управления и групп элементов управления"](#find-the-ids-of-controls-and-control-groups).

В следующем примере разметки группа элементов управления "Абзац Office" добавляется на настраиваемую вкладку и размещается сразу после настраиваемой группы.

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>Вставка встроенного элемента управления в пользовательскую группу

Чтобы вставить встроенный элемент управления Office в пользовательскую группу, добавьте [элемент OfficeControl](/javascript/api/manifest/group#officecontrol) в родительский элемент в качестве дочернего элемента **\<Group\>** . Атрибуту `id` элемента **\<OfficeControl\>** задается идентификатор встроенного элемента управления. См [. раздел "Поиск идентификаторов элементов управления и групп элементов управления"](#find-the-ids-of-controls-and-control-groups).

В следующем примере разметки элемент управления Надстрочный индекс Office добавляется в настраиваемую группу и размещается сразу после пользовательской кнопки.

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
> Пользователи могут настраивать ленту в приложении Office. Любые пользовательские настройки переопределяют параметры манифеста. Например, пользователь может удалить кнопку из любой группы и удалить любую группу со вкладки.

## <a name="find-the-ids-of-controls-and-control-groups"></a>Поиск идентификаторов элементов управления и групп элементов управления

Идентификаторы поддерживаемых элементов управления и групп управления находятся в файлах в идентификаторах элементов [управления Office в репозитории](https://github.com/OfficeDev/office-control-ids). Следуйте инструкциям в файле ReadMe этого репозитория.

## <a name="behavior-on-unsupported-platforms"></a>Поведение на неподдерживаемых платформах

Если надстройка установлена на платформе, которая не поддерживает набор обязательных элементов [AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets), то разметка, описанная в этой статье, игнорируется, а встроенные элементы управления и группы Office не будут отображаться в настраиваемых группах и вкладах. Чтобы запретить установку надстройки на платформах, не поддерживающих разметку, **\<Requirements\>** добавьте ссылку на набор обязательных элементов в разделе манифеста. Инструкции см. в [статье "Указание версий и платформ Office для размещения надстройки"](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in). Кроме того, можно спроектировать надстройку, чтобы иметь опыт, когда **AddinCommands 1.3** не поддерживается, как описано в разделе "Конструктор для альтернативных [возможностей"](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences). Например, если надстройка содержит инструкции, предполагающее, что встроенные кнопки находятся в пользовательских группах, можно разработать версию, предполагаемую, что встроенные кнопки находятся только в обычных местах.
