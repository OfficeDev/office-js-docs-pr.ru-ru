---
title: Элемент Method в файле манифеста
description: Элемент Method указывает отдельный метод из API Office JavaScript, который требуется Office надстройки для активации Office или переопределения параметров базового манифеста.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 052fb41a7077781843ea7e63d9601a819058dfa6
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222271"
---
# <a name="method-element"></a>Элемент Method

Значение этого элемента зависит от того, где он используется в манифесте.

## <a name="in-the-base-manifest"></a>В базовом манифесте

Когда используется в базовом манифесте  (то есть элемент требования к бабушке и дедушке является прямым ребенком [OfficeApp),](officeapp.md)элемент Method указывает отдельный метод из Office API JavaScript Office, необходимый для активации Office. 

**Тип надстройки:** контентные надстройки и надстройки области задач.

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>Как правнук элемента VersionOverrides

Указывает отдельный метод из API Office JavaScript, который должен поддерживаться версией и платформой Office (например, Windows, Mac, web и iOS или iPad) для того, чтобы [ВерсияOverrides](versionoverrides.md) вступила в силу.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- То же, что и элемент ["Требования к бабушке и дедушке".](requirements.md)

**Связанные с этими наборами требований:**

- То же, что и элемент ["Требования к бабушке и дедушке".](requirements.md)

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

## <a name="remarks"></a>Комментарии

Элементы **Methods** и **Method** не поддерживаются почтовыми надстройки при их использования в базовом манифесте. Дополнительные сведения о наборах обязательных элементов см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

> [!IMPORTANT]
> Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в Office [API JavaScript.](../../develop/understanding-the-javascript-api-for-office.md)
