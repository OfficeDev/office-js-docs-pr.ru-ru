---
title: Элемент SourceLocation в файле манифеста
description: Элемент SourceLocation указывает расположение исходных файлов для надстройки Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 447adb7df7d0c59305fe5046357959fcd7824735
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641405"
---
# <a name="sourcelocation-element"></a>Элемент SourceLocation

Указывает расположение исходных файлов для надстройки Office в виде URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Содержится в

- [DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)
- [FormSettings](formsettings.md) (почтовые надстройки)
- [ExtensionPoint](extensionpoint.md) (контекстные и лаунчевент (Предварительная версия) почтовые надстройки)

## <a name="can-contain"></a>Может содержать

[Override](override.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).|
