---
title: Элемент Set в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324808"
---
# <a name="set-element"></a>Элемент Set

Задает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Содержится в

[Sets](sets.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя [набора требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **дефаултминверсион**, если оно указано в элементе родительских [наборов](sets.md) .|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`. Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек). Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.
