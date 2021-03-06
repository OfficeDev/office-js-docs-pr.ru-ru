---
title: Работа с расширенными переопределениями манифеста
description: Узнайте, как настроить функции расширяемости с расширенными переопределениями манифеста.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505574"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a><span data-ttu-id="2ceeb-103">Работа с расширенными переопределениями манифеста</span><span class="sxs-lookup"><span data-stu-id="2ceeb-103">Work with Extended Overrides of the manifest</span></span>

<span data-ttu-id="2ceeb-104">Некоторые функции extensibility надстройки Office настроены с JSON-файлами, которые находятся на сервере, а не с XML-манифестом надстройки.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-104">Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="2ceeb-105">В этой статье предполагается, что вы знакомы с манифестами надстройки Office и их ролью в надстройки. Пожалуйста, [ознакомьтесь с XML-манифестом](add-in-manifests.md)надстройки Office, если вы еще не были недавно.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-105">This article assumes that you're familiar with Office add-in manifests and their role in add-ins. Please read [Office Add-ins XML manifest](add-in-manifests.md), if you haven't recently.</span></span>

<span data-ttu-id="2ceeb-106">В следующей таблице указаны функции расширяемости, которые требуют расширенного переопределения, а также ссылки на документацию по этой функции.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-106">The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.</span></span>

| <span data-ttu-id="2ceeb-107">Возможность</span><span class="sxs-lookup"><span data-stu-id="2ceeb-107">Feature</span></span> | <span data-ttu-id="2ceeb-108">Инструкции по разработке</span><span class="sxs-lookup"><span data-stu-id="2ceeb-108">Development Instructions</span></span> |
| :----- | :----- |
| <span data-ttu-id="2ceeb-109">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="2ceeb-109">Keyboard shortcuts</span></span> | [<span data-ttu-id="2ceeb-110">Добавление ярлыков настраиваемой клавиатуры в надстройки Office</span><span class="sxs-lookup"><span data-stu-id="2ceeb-110">Add Custom keyboard shortcuts to your Office Add-ins</span></span>](../design/keyboard-shortcuts.md) |

<span data-ttu-id="2ceeb-111">Схема, определяемая форматом JSON, имеет [схему расширенного манифеста.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)</span><span class="sxs-lookup"><span data-stu-id="2ceeb-111">The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!TIP]
> <span data-ttu-id="2ceeb-112">Эта статья несколько абстрактна.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-112">This article is somewhat abstract.</span></span> <span data-ttu-id="2ceeb-113">Рассмотрите чтение одной из статей в таблице, чтобы добавить ясность к понятиям.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-113">Consider reading one of the articles in the table to add clarity to the concepts.</span></span>

## <a name="tell-office-where-to-find-the-json-file"></a><span data-ttu-id="2ceeb-114">Скажите Office, где найти файл JSON</span><span class="sxs-lookup"><span data-stu-id="2ceeb-114">Tell Office where to find the JSON file</span></span>

<span data-ttu-id="2ceeb-115">Используйте манифест, чтобы сообщить Office, где найти файл JSON.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-115">Use the manifest to tell Office where to find the JSON file.</span></span> <span data-ttu-id="2ceeb-116">Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="2ceeb-116">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="2ceeb-117">Установите атрибут `Url` для полного URL-адреса файла JSON.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-117">Set the `Url` attribute to the full URL of a JSON file.</span></span> <span data-ttu-id="2ceeb-118">Ниже приводится пример простейшего `<ExtendedOverrides>` элемента.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-118">The following is an example of the simplest possible `<ExtendedOverrides>` element.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="2ceeb-119">Ниже приводится пример очень простого расширенного переопределения JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-119">The following is an example of a very simple extended overrides JSON file.</span></span> <span data-ttu-id="2ceeb-120">Он назначает клавишу ярлык CTRL+SHIFT+A функции (определенной в другом месте), которая открывает области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-120">It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a><span data-ttu-id="2ceeb-121">Локализация расширенного переопределения файла</span><span class="sxs-lookup"><span data-stu-id="2ceeb-121">Localize the extended overrides file</span></span>

<span data-ttu-id="2ceeb-122">Если надстройка поддерживает несколько локальных элементов, можно использовать атрибут элемента, чтобы указать `ResourceUrl` `<ExtendedOverrides>` Office на файл локализованных ресурсов.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-122">If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the `<ExtendedOverrides>` element to point Office to a file of localized resources.</span></span> <span data-ttu-id="2ceeb-123">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="2ceeb-123">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="2ceeb-124">Дополнительные сведения о том, как создавать и использовать файл ресурсов, как ссылаться на его ресурсы в расширенном файле переопределения, а также дополнительные параметры, не рассмотренные здесь, см. в материале [Localize extended overrides.](localization.md#localize-extended-overrides)</span><span class="sxs-lookup"><span data-stu-id="2ceeb-124">For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).</span></span>
