---
title: Элемент Version в файле манифеста
description: Элемент Version указывает версию надстройки Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173936"
---
# <a name="version-element"></a>Элемент Version

Указывает версию надстройки Office. Номер версии: 1, 2, 3 или 4 части (например, n, n.n, n.n.n или n.n.n).

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Замечания

Каждая часть номера версии может быть не более 5 цифр.
