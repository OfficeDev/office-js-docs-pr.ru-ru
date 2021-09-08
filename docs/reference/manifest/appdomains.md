---
title: Элемент AppDomains в файле манифеста
description: Перечислены все домены в дополнение к домену, указанному в элементе, Office надстройка будет использовать и должна доверяться `SourceLocation` Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936409"
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
