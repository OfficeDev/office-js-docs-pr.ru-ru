---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из API JavaScript для Office, необходимый для активации надстроек Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c3531475a920fd24ce8390170b5f4728d4dcd0e0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611759"
---
# <a name="method-element"></a>Элемент Method

Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.

**Тип надстройки:** контентные надстройки и надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Содержится в

[Методы](methods.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Примечания

`Methods`Элементы и `Method` не поддерживаются почтовыми надстройками. Дополнительные сведения о наборах требований: [версии и наборы](../../develop/office-versions-and-requirement-sets.md)обязательных элементов для Office.

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, можно узнать в статье Общие сведения об [API JavaScript для Office](../../develop/understanding-the-javascript-api-for-office.md).
