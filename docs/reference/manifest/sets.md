---
title: Элемент Sets в файле манифеста
description: Элемент Sets указывает минимальный набор API Office JavaScript, необходимый Office надстройки для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937479"
---
# <a name="sets-element"></a>Элемент Sets

Указывает минимальный подмножество API JavaScript Office, который требуется Office надстройки для активации.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Содержится в

[Requirements](requirements.md)

## <a name="can-contain"></a>Может содержать

[Set](set.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Указывает значение атрибута **MinVersion** по умолчанию для всех [элементов](set.md) набора детей. Значение по умолчанию: "1.1".|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Дополнительные сведения о атрибуте **MinVersion** элемента **Set** и **атрибуте DefaultMinVersion** элемента **Sets** см. в элементе [Set the Requirements in the manifest.](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)

