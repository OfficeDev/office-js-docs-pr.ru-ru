---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337197"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

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

Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).
