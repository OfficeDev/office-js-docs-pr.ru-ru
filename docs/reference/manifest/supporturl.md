---
title: Элемент SupportUrl в файле манифеста
description: Элемент SupportUrl указывает URL-адрес страницы, которая предоставляет сведения о поддержке вашей надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1d76afeaaceafc9e8786070338d69cea1b73635d20cd5a729d7e3d859b952494
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096363"
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

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL-адрес|Обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|
