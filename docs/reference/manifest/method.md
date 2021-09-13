---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из Office API JavaScript, который требуется Office надстройки для активации.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 037446f5027a97214d2b1be6ee99c8f6822b33b9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151245"
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

## <a name="remarks"></a>Примечания

Эти `Methods` элементы и элементы не `Method` поддерживаются почтовыми надстройки. Дополнительные сведения о наборах требований [см. в Office версиях и наборах требований.](../../develop/office-versions-and-requirement-sets.md)

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в Office [API JavaScript.](../../develop/understanding-the-javascript-api-for-office.md)
