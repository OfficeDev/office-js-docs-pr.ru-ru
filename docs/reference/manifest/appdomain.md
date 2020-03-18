---
title: Элемент AppDomain в файле манифеста
description: Задает дополнительные домены, которые загружают страницы в окне надстройки.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718456"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Задает дополнительные домены, которые загружают страницы в окне надстройки. Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain</AppDomain>`).
> 2. *Не* ставьте закрывающую косую черту ("/") для значения.

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Примечания

Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).
