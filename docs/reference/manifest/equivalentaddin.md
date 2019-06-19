---
title: Элемент Екуивалентаддин в файле манифеста
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059925"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="b3978-102">Элемент Екуивалентаддин</span><span class="sxs-lookup"><span data-stu-id="b3978-102">EquivalentAddin element</span></span>

<span data-ttu-id="b3978-103">Задает обратную совместимость для эквивалентной надстройки COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="b3978-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="b3978-104">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="b3978-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="b3978-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b3978-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="b3978-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b3978-106">Contained in</span></span>

[<span data-ttu-id="b3978-107">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="b3978-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="b3978-108">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="b3978-108">Must contain</span></span>

[<span data-ttu-id="b3978-109">Тип</span><span class="sxs-lookup"><span data-stu-id="b3978-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="b3978-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="b3978-110">Can contain</span></span>

<span data-ttu-id="b3978-111">[](progid.md)
[Имя файла](filename.md) ProgID</span><span class="sxs-lookup"><span data-stu-id="b3978-111">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="b3978-112">Примечания</span><span class="sxs-lookup"><span data-stu-id="b3978-112">Remarks</span></span>

<span data-ttu-id="b3978-113">Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgId` `Type` элемента:.</span><span class="sxs-lookup"><span data-stu-id="b3978-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="b3978-114">Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:</span><span class="sxs-lookup"><span data-stu-id="b3978-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="b3978-115">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="b3978-115">See also</span></span>

- [<span data-ttu-id="b3978-116">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="b3978-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="b3978-117">Обеспечение совместимости надстройки Excel с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="b3978-117">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)