---
title: Элемент SupportUrl в файле манифеста
description: Элемент SupportUrl указывает URL-адрес страницы, предоставляющей сведения о поддержке надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f75ee811699823a501ac594e66daaaf3f93c2782
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608707"
---
# <a name="supporturl-element"></a>Элемент SupportUrl

Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.

## <a name="syntax"></a>Синтаксис

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|  Элемент | Обязательный | Описание  |
|:-----|:-----|:-----|
|  [Override](override.md)   | Нет | Задает параметр для URL-адресов дополнительных языковых стандартов |

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|
