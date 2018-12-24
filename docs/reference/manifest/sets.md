---
title: Элемент Sets в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433658"
---
# <a name="sets-element"></a>Элемент Sets

Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.

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

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](set.md). Значение по умолчанию: "1.1".|

## <a name="remarks"></a>Примечания

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

