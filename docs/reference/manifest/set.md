---
title: Элемент Set в файле манифеста
description: Элемент Set указывает Office API JavaScript, заданный Office надстройки для активации.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 93524d64fd915d6f42f4e4a0cd0ab6cc3335f4ce
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154464"
---
# <a name="set-element"></a>Элемент Set

Указывает набор требований из Office API JavaScript, который требуется Office надстройки для активации.

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
|MinVersion|string|необязательный|Указывает минимальную версию набора API, необходимую надстройке. Переопределяет значение **DefaultMinVersion,** если оно указано в элементе родительских [наборов.](sets.md)|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в элементе [Set the Requirements in the manifest.](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)

> [!IMPORTANT]
> Для почтовых надстроек доступен только один набор обязательных элементов `"Mailbox"`. Он содержит все подмножество API, поддерживаемое почтовыми надстройками Outlook, а в манифесте почтовой надстройки необходимо указать набор обязательных элементов `"Mailbox"` (это обязательно для почтовых надстроек, в отличие от надстроек области задачи и контентных надстроек). Кроме того, в почтовых надстройках невозможно объявить поддержку определенных методов.
