---
title: Элемент AppDomains в файле манифеста
description: Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, который надстройка Office будет использовать для загрузки страниц.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: f60579d773e81a7e8006bafcf1c151874af42aeb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720703"
---
# <a name="appdomains-element"></a>Элемент AppDomains

Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, который надстройка Office будет использовать для загрузки страниц. Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке. Для каждого дополнительного домена укажите элемент AppDomain.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

[AppDomain](appdomain.md)

## <a name="remarks"></a>Примечания

По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md). Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**. Этот элемент не может быть пустым.
