---
title: Элемент OverriddenByRibbonApi в файле манифеста
description: Узнайте, как указать, что настраиваемая вкладка, группа, элемент управления или пункт меню не должны отображаться, когда она также является частью настраиваемой контекстной вкладки.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 62aa484057221f9cd7f41af9c8b9210cdb5b3376
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50174002"
---
# <a name="overriddenbyribbonapi-element"></a><span data-ttu-id="edac1-103">Элемент OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="edac1-103">OverriddenByRibbonApi element</span></span>

<span data-ttu-id="edac1-104">Указывает, будет ли элемент [customTab,](customtab.md) [](control.md#menu-dropdown-button-controls) [group,](group.md) [button,](control.md#button-control) menu или элемент меню скрытыми в сочетаниях приложений и платформ, которые поддерживают API[(Office.ribbon.requestCreateControls),](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)устанавливая настраиваемые контекстные вкладки на ленте.</span><span class="sxs-lookup"><span data-stu-id="edac1-104">Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.</span></span>

<span data-ttu-id="edac1-105">Если он опущен, значение по `false` умолчанию: .</span><span class="sxs-lookup"><span data-stu-id="edac1-105">If it is omitted, the default is `false`.</span></span> <span data-ttu-id="edac1-106">Если он используется, он должен быть первым *родительским* элементом.</span><span class="sxs-lookup"><span data-stu-id="edac1-106">If it is used, it must be the *first* child element of its parent element.</span></span>

> [!NOTE]
> <span data-ttu-id="edac1-107">Чтобы полностью понять этот элемент, прочитайте статью ["Реализация альтернативного интерфейса](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)пользователя, если настраиваемые контекстные вкладки не поддерживаются".</span><span class="sxs-lookup"><span data-stu-id="edac1-107">For a full understanding of this element, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

<span data-ttu-id="edac1-108">Цель этого элемента — создать в надстройке возможность отката, которая реализует настраиваемые контекстные вкладки, когда надстройка работает в приложении или платформе, которые не поддерживают настраиваемые контекстные вкладки.</span><span class="sxs-lookup"><span data-stu-id="edac1-108">The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> <span data-ttu-id="edac1-109">Основная стратегия заключается в том, что вы дублируете некоторые или все группы и элементы управления  с настраиваемой контекстной вкладки на одну или несколько настраиваемой основной вкладки (то есть неконтекстуальные настраиваемые вкладки).</span><span class="sxs-lookup"><span data-stu-id="edac1-109">The essential strategy is that you duplicate some or all of the groups and controls from your custom contextual tab onto one or more custom core tabs (that is, *noncontextual* custom tabs).</span></span> <span data-ttu-id="edac1-110">Затем, чтобы эти группы и элементы управления появлялись, когда настраиваемые контекстные вкладки  не поддерживаются, но не отображаются при поддержке настраиваемой контекстной вкладки, необходимо добавить в качестве первого потомка элементов  `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` **CustomTab,** **Group,** **Control** или menu **Item.**</span><span class="sxs-lookup"><span data-stu-id="edac1-110">Then, to ensure that these groups and controls appear when custom contextual tabs are *not* supported, but do not appear when custom contextual tabs *are* supported, you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**, **Group**, **Control**, or menu **Item** elements.</span></span> <span data-ttu-id="edac1-111">Это может быть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="edac1-111">The effect of doing so is the following:</span></span>

- <span data-ttu-id="edac1-112">Если надстройка работает на приложениях и платформах, поддерживаюх настраиваемые контекстные вкладки, дублированные вкладки, группы и элементы управления не будут отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="edac1-112">If the add-in runs on an application and platform that support custom contextual tabs, then the duplicated tabs, groups, and controls won't appear on the ribbon.</span></span> <span data-ttu-id="edac1-113">Вместо этого настраиваемая контекстная вкладка будет установлена, когда надстройка вызывает `requestCreateControls` метод.</span><span class="sxs-lookup"><span data-stu-id="edac1-113">Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="edac1-114">Если надстройка работает в приложении или платформе, не поддерживаю которой настраиваемые контекстные вкладки, дублированные вкладки, группы и элементы управления будут отображаться на ленте. </span><span class="sxs-lookup"><span data-stu-id="edac1-114">If the add-in runs on an application or platform that *doesn't* support custom contextual tabs, then the duplicated tabs, groups, and controls will appear on the ribbon.</span></span>

## <a name="examples"></a><span data-ttu-id="edac1-115">Примеры</span><span class="sxs-lookup"><span data-stu-id="edac1-115">Examples</span></span>

### <a name="overriding-an-entire-tab"></a><span data-ttu-id="edac1-116">Переопределение всей вкладки</span><span class="sxs-lookup"><span data-stu-id="edac1-116">Overriding an entire tab</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-group"></a><span data-ttu-id="edac1-117">Переопределение группы</span><span class="sxs-lookup"><span data-stu-id="edac1-117">Overriding a group</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Control  xsi:type="Button" id="MyButton">
        <!-- Child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-control"></a><span data-ttu-id="edac1-118">Переопределение управления</span><span class="sxs-lookup"><span data-stu-id="edac1-118">Overriding a control</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
        <!-- Other child elements omitted. -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```

### <a name="overriding-a-menu-item"></a><span data-ttu-id="edac1-119">Переопределение элемента меню</span><span class="sxs-lookup"><span data-stu-id="edac1-119">Overriding a menu item</span></span>


```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Menu" id="MyMenu">
        <!-- Other child elements omitted. -->
        <Items>
          <Item id="showGallery">
            <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
            <!-- Other child elements omitted. -->
          </Item>
        </Items>
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
