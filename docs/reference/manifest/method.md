---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из Office API JavaScript, который требуется Office надстройки для активации.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 811cd84e1ad2aade8b7042eefa822eee6b2ab200a8fa1b71c9fe5fc34874ec66
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089732"
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
