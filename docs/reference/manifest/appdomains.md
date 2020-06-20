---
title: Элемент AppDomains в файле манифеста
description: Список всех доменов в дополнение к домену, указанному в `SourceLocation` элементе, который будет использоваться вашей надстройкой Office и должен быть доверенным для Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778657"
---
# <a name="appdomains-element"></a>Элемент AppDomains

Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, что ваша надстройка Office будет использовать и должна быть доверенной для Office. Это позволяет страницам в доменах совершать вызовы Office.js API из IFrames в надстройке и имеет другие эффекты. Для каждого дополнительного домена укажите элемент **AppDomain**.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Существуют ограничения на то, что может быть значением элемента **AppDomain** . Дополнительные сведения см. в разделе [AppDomain](appdomain.md).

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

[AppDomain](appdomain.md)

## <a name="remarks"></a>Примечания

По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md). Этот элемент не может быть пустым.
