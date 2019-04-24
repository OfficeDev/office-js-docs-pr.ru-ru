---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451978"
---
# <a name="sourcelocation-element"></a>Элемент SourceLocation

Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Содержится в

- [DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)
- [FormSettings](formsettings.md) (почтовые надстройки)
- [ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)

## <a name="can-contain"></a>Может содержать

[Override](override.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).|
