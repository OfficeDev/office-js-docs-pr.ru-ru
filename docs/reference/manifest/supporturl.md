---
title: Элемент SupportUrl в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 00234ef9fe8960b9956e6a2595e2e2e71bfb97c6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432671"
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
