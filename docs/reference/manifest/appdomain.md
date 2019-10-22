---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575641"
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

Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).
