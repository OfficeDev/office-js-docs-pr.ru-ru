---
title: Элемент AppDomains в файле манифеста
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433683"
---
# <a name="appdomains-element"></a>Элемент AppDomains

Определяет все домены, кроме указанного в элементе SourceLocation, которые надстройка Office будет использовать для загрузки страниц. Для каждого дополнительного домена укажите элемент AppDomain.

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
