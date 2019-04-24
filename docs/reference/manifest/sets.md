---
title: Элемент Sets в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450424"
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

Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

