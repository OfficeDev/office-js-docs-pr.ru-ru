---
title: Элемент OfficeTab в файле манифеста
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 1bf9f1d1e08a8147b52f93923229ef8fb8556fcf
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952273"
---
# <a name="officetab-element"></a>Элемент OfficeTab

Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка. Этот элемент обязательный.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  Группа      | Да |  Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.  |

Ниже приведены допустимые значения `id` для вкладок каждого ведущего приложения. Значения, **выделенные жирным шрифтом** , поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или более поздней версии в Windows и Word Online).

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
