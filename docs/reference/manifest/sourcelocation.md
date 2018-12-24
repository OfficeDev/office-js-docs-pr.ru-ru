---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433518"
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

[Переопределение](override.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).|
