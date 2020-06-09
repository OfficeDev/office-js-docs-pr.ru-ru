---
title: Элемент Екуивалентаддин в файле манифеста
description: Задает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611563"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="9a48f-103">Элемент Екуивалентаддин</span><span class="sxs-lookup"><span data-stu-id="9a48f-103">EquivalentAddin element</span></span>

<span data-ttu-id="9a48f-104">Задает обратную совместимость для эквивалентной надстройки COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="9a48f-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="9a48f-105">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="9a48f-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="9a48f-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9a48f-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="9a48f-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9a48f-107">Contained in</span></span>

[<span data-ttu-id="9a48f-108">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="9a48f-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="9a48f-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="9a48f-109">Must contain</span></span>

[<span data-ttu-id="9a48f-110">Тип</span><span class="sxs-lookup"><span data-stu-id="9a48f-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="9a48f-111">Может содержать</span><span class="sxs-lookup"><span data-stu-id="9a48f-111">Can contain</span></span>

<span data-ttu-id="9a48f-112">[ProgID](progid.md) 
 [Имя файла](filename.md)</span><span class="sxs-lookup"><span data-stu-id="9a48f-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="9a48f-113">Примечания</span><span class="sxs-lookup"><span data-stu-id="9a48f-113">Remarks</span></span>

<span data-ttu-id="9a48f-114">Чтобы указать надстройку COM в качестве эквивалентной надстройки, укажите оба `ProgId` `Type` элемента:.</span><span class="sxs-lookup"><span data-stu-id="9a48f-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="9a48f-115">Чтобы указать XLL в качестве эквивалентной надстройки, укажите оба `FileName` `Type` элемента:</span><span class="sxs-lookup"><span data-stu-id="9a48f-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="9a48f-116">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="9a48f-116">See also</span></span>

- [<span data-ttu-id="9a48f-117">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="9a48f-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="9a48f-118">Обеспечение совместимости надстройки Excel с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="9a48f-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)