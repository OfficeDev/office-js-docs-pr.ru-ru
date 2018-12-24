---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433070"
---
# <a name="appdomain-element"></a>Элемент AppDomain

Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.

## <a name="syntax"></a>Синтаксис

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Содержится в

[AppDomains](appdomains.md)

## <a name="remarks"></a>Примечания

Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).
