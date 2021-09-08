---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из Office API JavaScript, который требуется Office надстройки для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936546"
---
# <a name="method-element"></a>Элемент Method

Указывает отдельный метод из API Office JavaScript, который требуется Office надстройки для активации.

**Тип надстройки:** контентные надстройки и надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Содержится в

[Методы](methods.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Замечания

Эти `Methods` элементы и элементы не `Method` поддерживаются почтовыми надстройки. Дополнительные сведения о наборах требований [см. в Office версиях и наборах требований.](../../develop/office-versions-and-requirement-sets.md)

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в Office [API JavaScript.](../../develop/understanding-the-javascript-api-for-office.md)
