---
title: Элемент Version в файле манифеста
description: Элемент Version указывает Office надстройки.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 9641153cbe6fa0284986b8dd286ba2114b32a82894bd5f8d33516e2a56c90be9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096333"
---
# <a name="version-element"></a>Элемент Version

Указывает версию надстройки Office. Номер версии может быть 1, 2, 3 или 4 частей (например, n, n.n, n.n.n или n.n.n.n.).

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Каждая часть номера версии может быть не более 5 цифр.
