---
title: Элемент AppDomains в файле манифеста
description: Перечислены все домены в дополнение к домену, указанному в элементе, Office надстройка будет использовать и должна доверяться `SourceLocation` Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 55401d62e88cc1f2d67d13de0997a40db7a3f6b0c2f8997aa1b976962c8c797f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096536"
---
# <a name="appdomains-element"></a>Элемент AppDomains

Перечислены все домены, в дополнение к домену, указанному в элементе, который будет Office надстройка и которому следует доверять `SourceLocation` Office. Это позволяет страницам в доменах звонить Office.js API из IFrames в надстройки и имеет другие эффекты. Для каждого дополнительного домена укажите элемент **AppDomain**.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Существуют ограничения по значению элемента **AppDomain.** Дополнительные сведения см. в [приложении AppDomain.](appdomain.md)

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

[AppDomain](appdomain.md)

## <a name="remarks"></a>Примечания

По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md). Этот элемент не может быть пустым.
