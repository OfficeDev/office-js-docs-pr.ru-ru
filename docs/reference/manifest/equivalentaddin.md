---
title: Элемент Екуивалентаддин в файле манифеста
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356912"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="e7c09-102">Элемент Екуивалентаддин</span><span class="sxs-lookup"><span data-stu-id="e7c09-102">EquivalentAddin element</span></span>

<span data-ttu-id="e7c09-103">Задает обратную совместимость для эквивалентной надстройки COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="e7c09-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="e7c09-104">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="e7c09-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="e7c09-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e7c09-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="e7c09-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e7c09-106">Contained in</span></span>

[<span data-ttu-id="e7c09-107">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="e7c09-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="e7c09-108">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="e7c09-108">Must contain</span></span>

[<span data-ttu-id="e7c09-109">Type</span><span class="sxs-lookup"><span data-stu-id="e7c09-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="e7c09-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="e7c09-110">Can contain</span></span>

<span data-ttu-id="e7c09-111">[](progid.md)
[Имя файла](filename.md) ProgID</span><span class="sxs-lookup"><span data-stu-id="e7c09-111">[ProgID](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="e7c09-112">Примечания</span><span class="sxs-lookup"><span data-stu-id="e7c09-112">Remarks</span></span>

<span data-ttu-id="e7c09-113">Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgID` `Type` элемента:.</span><span class="sxs-lookup"><span data-stu-id="e7c09-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements.</span></span> <span data-ttu-id="e7c09-114">Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:</span><span class="sxs-lookup"><span data-stu-id="e7c09-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="e7c09-115">См. также</span><span class="sxs-lookup"><span data-stu-id="e7c09-115">See also</span></span>

- [<span data-ttu-id="e7c09-116">Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL</span><span class="sxs-lookup"><span data-stu-id="e7c09-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="e7c09-117">Обеспечение совместимости надстройки Office с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="e7c09-117">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)