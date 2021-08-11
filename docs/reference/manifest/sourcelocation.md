---
title: Элемент SourceLocation в файле манифеста
description: Элемент SourceLocation указывает расположение исходных файлов для Office надстройки.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 6830a26cf192802c97c486511695b4ace35ac8263cfcd30ceaf71398f0d83a07
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095011"
---
# <a name="sourcelocation-element"></a>Элемент SourceLocation

Указывает расположение исходных файлов для надстройки Office как URL-адрес длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Содержится в

- [DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)
- [FormSettings](formsettings.md) (почтовые надстройки)
- [ExtensionPoint](extensionpoint.md) (надстройки для почты Contextual и LaunchEvent)

## <a name="can-contain"></a>Может содержать

[Override](override.md)

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).|
