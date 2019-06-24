---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d073d712cec2fd58e957ffe8f344d7443d1e896e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127564"
---
# <a name="officetab-element"></a>Элемент OfficeTab

Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  Группа      | Да |  Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.  |

Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения. Значения, **выделенные полужирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word в Интернете).

### <a name="outlook"></a>Outlook

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval 

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

Группа точек расширения пользовательского интерфейса на вкладке. В группе может быть до шести элементов управления. Атрибут **id** обязательный, и каждый атрибут **id** должен быть уникальным в манифесте. Атрибут **id** — это строка длиной до 125 символов. См. статью об[элементе Group](group.md).

## <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
