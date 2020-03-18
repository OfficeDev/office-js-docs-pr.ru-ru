---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из API JavaScript для Office, необходимый для активации надстроек Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 5da25616d25a8d7454fc847727cda38a9935b5c7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720584"
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
|Имя|string|Обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы указать `getSelectedDataAsync` метод, необходимо указать. `"Document.getSelectedDataAsync"`|

## <a name="remarks"></a>Примечания

Элементы `Methods` и `Method` не поддерживаются почтовыми надстройками. Дополнительные сведения о наборах требований: [версии и наборы](../../develop/office-versions-and-requirement-sets.md)обязательных элементов для Office.

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, можно узнать в статье Общие сведения об [API JavaScript для Office](../../develop/understanding-the-javascript-api-for-office.md).
