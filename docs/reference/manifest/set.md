---
title: Элемент Set в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0f408d698d297eaa6287ff268bdb7fc737a5a24d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452034"
---
# <a name="set-element"></a>Элемент Set

Указывает набор требований из API JavaScript для Office, необходимый для активации надстройки Office.

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
|Имя|строка|Обязательный|Имя [набора требований](/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion**, если оно указано в родительском элементе [Sets](sets.md).|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`. Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек). Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.
