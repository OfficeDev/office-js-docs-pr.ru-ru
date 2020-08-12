---
title: Элемент Set в файле манифеста
description: Элемент Set указывает набор обязательных элементов API JavaScript для Office, необходимый для активации надстройки Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641419"
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

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя [набора требований](../../develop/office-versions-and-requirement-sets.md).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **дефаултминверсион**, если оно указано в элементе родительских [наборов](sets.md) .|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **дефаултминверсион** элемента **Sets** приведены в разделе [set the требований в манифесте](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT]
> Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`. Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек). Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.
