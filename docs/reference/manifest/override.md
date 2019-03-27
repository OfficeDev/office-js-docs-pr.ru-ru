---
title: Элемент Override в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870060"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ указать значение параметра для дополнительного языкового стандарта.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Содержится в

|**Element**|
|:-----|
|[CitationText](citationtext.md)|
|[Описание](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Языковой стандарт|string|Обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.|
|Значение|string|Обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

## <a name="see-also"></a>См. также

- [Локализация надстроек для Office](/office/dev/add-ins/develop/localization)
    
