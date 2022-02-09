---
title: Элемент Control в файле манифеста
description: Определяет управление, которое выполняет действие или запускает области задач.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa7ff9b0162070b378352ce187de15a34323b998
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467838"
---
# <a name="control-element"></a>Элемент Control

Определяет управление, которое выполняет действие или запускает области задач. Элемент **Control** может быть кнопкой или пунктом меню. Элемент **Group** должен содержать по крайней мере один элемент [Control](group.md).

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Для надстройки области задач.)
- Некоторые детские элементы могут быть связаны с дополнительными наборами требований.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|**xsi:type**|Да|Тип определяемого элемента управления. Может быть `Button`, `Menu`или `MobileButton`. |
|**id**|Да|ИД элемента управления. Может содержать до 125 знаков. Должно быть уникальным для всех **элементов управления** в манифесте.|

> [!NOTE]
> Значение `MobileButton` для **xsi:type** определено в схеме 1.1 VersionOverrides. Применяется только к элементам **Control**, которые содержатся в элементе [MobileFormFactor](mobileformfactor.md).

## <a name="child-elements"></a>Дочерние элементы

Допустимые элементы ребенка зависят от значения **атрибута xsi:type** .

- [Тип элемента Control button](control-button.md)
- [Элемент Control типа меню](control-menu.md)
- [Элемент Управления типа MobileButton](control-mobilebutton.md)
