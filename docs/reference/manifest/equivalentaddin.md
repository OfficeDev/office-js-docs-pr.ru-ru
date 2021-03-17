---
title: Элемент EquivalentAddin в файле манифеста
description: Указывает обратную совместимость для эквивалентной надстройки COM или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836839"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="eadc9-103">Элемент EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="eadc9-103">EquivalentAddin element</span></span>

<span data-ttu-id="eadc9-104">Указывает обратную совместимость для эквивалентной надстройки COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="eadc9-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="eadc9-105">**Тип надстройки:** Области задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="eadc9-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="eadc9-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="eadc9-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="eadc9-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="eadc9-107">Contained in</span></span>

[<span data-ttu-id="eadc9-108">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="eadc9-108">EquivalentAddins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="eadc9-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="eadc9-109">Must contain</span></span>

[<span data-ttu-id="eadc9-110">Тип</span><span class="sxs-lookup"><span data-stu-id="eadc9-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="eadc9-111">Может содержать</span><span class="sxs-lookup"><span data-stu-id="eadc9-111">Can contain</span></span>

<span data-ttu-id="eadc9-112">[ProgId](progid.md) 
 [FileName](filename.md)</span><span class="sxs-lookup"><span data-stu-id="eadc9-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="eadc9-113">Примечания</span><span class="sxs-lookup"><span data-stu-id="eadc9-113">Remarks</span></span>

<span data-ttu-id="eadc9-114">Чтобы указать надстройки COM в качестве эквивалентной надстройки, укажите как элементы, так `ProgId` `Type` и элементы.</span><span class="sxs-lookup"><span data-stu-id="eadc9-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="eadc9-115">Чтобы указать XLL в качестве эквивалентной надстройки, укажите как элементы, так `FileName` `Type` и элементы.</span><span class="sxs-lookup"><span data-stu-id="eadc9-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="eadc9-116">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="eadc9-116">See also</span></span>

- [<span data-ttu-id="eadc9-117">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="eadc9-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="eadc9-118">Убедитесь, что надстройка Office совместима с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="eadc9-118">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)