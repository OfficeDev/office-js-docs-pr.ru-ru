---
title: Элемент Version в файле манифеста
description: Элемент Version указывает Office надстройки.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939324"
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
