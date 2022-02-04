---
title: Элемент OfficeTab в файле манифеста
description: 'Элемент OfficeTab определяет вкладку ленты, на которой отображается команда надстройки.'
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="officetab-element"></a>Элемент OfficeTab

Определяет вкладку ленты, на которой отображается команда надстройки. Это может быть вкладка по умолчанию (**главная****, сообщение** или **собрание) или** настраиваемая вкладка, определяемая надстройками. Этот элемент обязательный.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Taskpane 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  Группа      | Да |  Определяет группу команд. На вкладке по умолчанию можно добавить только одну группу для каждой надстройки.  |

Ниже по приложению допустимы `id` значения вкладок. Значения **жирным шрифтом** поддерживаются как на рабочем столе, так и в Интернете (например, Word 2016 или Windows и Word в Интернете).

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

Группа точек расширения пользовательского интерфейса на вкладке. Группа может иметь до шести элементов управления. Атрибут **id** необходим, и каждый **id** должен быть уникальным для всех групп манифеста. **ID —** это строка с максимальным значением 125 символов. См [. элемент Group](group.md).

## <a name="officetab-example"></a>Пример элемента OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="Contoso.msgreadTabMessage.group1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
