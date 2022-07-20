---
title: Карточка значения сущности типов данных API JavaScript для Excel
description: Узнайте, как использовать карточки значений сущностей с типами данных в надстройке Excel.
ms.date: 07/14/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7eb6251467b73af5e592d4cf013e899207944192
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889168"
---
# <a name="use-cards-with-entity-value-data-types-preview"></a>Использование карточек с типами данных значения сущности (предварительная версия)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

В этой статье описывается, как использовать [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) для создания модальных окон карточек в пользовательском интерфейсе Excel с типами данных значения сущности. Эти карточки могут отображать дополнительные сведения, содержащиеся в значении сущности, за пределами того, что уже отображается в ячейке, например связанные изображения, сведения о категории продукта и атрибуты данных.

Значение сущности , [или EntityCellValue](/javascript/api/excel/excel.entitycellvalue), является контейнером для типов данных и похож на объект в объектно-ориентированном программировании. В этой статье показано, как использовать свойства карточки значения сущности, параметры макета и функции атрибутов данных для создания значений сущностей, которые отображаются в виде карточек.

На следующем снимке экрана показан пример открытой карточки значения сущности, в данном случае для продукта **Tofu** из списка продуктов магазина продуктов.

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="Снимок экрана: тип данных значения сущности с окном карточки.":::

## <a name="card-properties"></a>Свойства карточки

Свойство значения сущности [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) позволяет задавать настраиваемые сведения о типах данных. Ключ `properties` принимает вложенные типы данных. Каждое вложенное свойство или тип данных должно иметь параметр `type` и параметр `basicValue` .

> [!IMPORTANT]
> Вложенные `properties` типы данных используются в сочетании со значениями [](#card-layout) макета карточки, описанными в следующем разделе статьи. После определения вложенного типа `properties``layouts` данных он должен быть назначен в свойстве для отображения на карточке.

В следующем фрагменте кода показан JSON для значения сущности с несколькими типами данных, вложенными в .`properties`

> [!NOTE]
> Чтобы узнать, как использовать этот JSON в полный пример кода, посетите репозиторий [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) .

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Image": {
            type: Excel.CellValueType.webImage,
            address: product.productImage || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

На следующем снимке экрана показана карточка значения сущности, которая использует предыдущий фрагмент кода. На снимке **экрана показаны сведения** об  идентификаторе **продукта, имени** **продукта,** изображении **, количестве** на единицу и ценах за единицу из предыдущего фрагмента кода.

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="Снимок экрана: тип данных значения сущности с окном макета карточки. На карточке отображаются имя продукта, идентификатор продукта, количество на единицу и сведения о ценах за единицу.":::

## <a name="card-layout"></a>Макет карточки

Свойство значения [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) [`card`](/javascript/api/excel/excel.entityviewlayouts) сущности создает объект a для сущности, а затем задает внешний вид этой карточки, например название карточки, изображение карточки и количество отображаемого раздела.

> [!IMPORTANT]
> Вложенные `layouts` значения используются в сочетании с типами данных [](#card-properties) свойств карточки, описанными в предыдущем разделе статьи. Чтобы его `properties` `layouts` можно было назначить для отображения на карточке, необходимо определить вложенный тип данных.

В свойстве `card` используйте объект [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) для определения компонентов `title`карточки, таких как , и `subTitle``sections`.

В следующем фрагменте кода JSON `card` `title` `mainImage` значения сущности показан макет с вложенными объектами и тремя объектами в `sections` карточке. Обратите внимание, что `title` свойство имеет `"Product Name"` соответствующий тип данных в предыдущем разделе статьи [о свойствах](#card-properties) карточки. Свойство `mainImage` также имеет соответствующий `"Image"` тип данных в предыдущем разделе. Свойство `sections` принимает вложенный массив и использует объект [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) для определения внешнего вида каждого раздела.

В каждом разделе карточки можно указать такие элементы `layout`, как , `title`и `properties`. Ключ `layout` использует объект [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) и принимает значение `"List"`. Ключ `properties` принимает массив строк. Обратите внимание, что `properties` значения, `"Product ID"`например, имеют соответствующие типы данных в предыдущем разделе статьи [о свойствах](#card-properties) карточки. Разделы также можно сворачивать и определять логическими значениями как свернутые или не свернутые при открытии карточки сущности в пользовательском интерфейсе Excel.

> [!NOTE]
> Чтобы узнать, как использовать этот JSON в полный пример кода, посетите репозиторий [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) .

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
            mainImage: { 
                property: "Image" 
            },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

На следующем снимке экрана показана карточка значения сущности, которая использует приведенные выше фрагменты кода. На снимке экрана показан `mainImage` объект в верхней части, `title` за которым следует объект, который использует название **продукта** и имеет значение **Tofu**. На снимке экрана также показано `sections`. Раздел **"Количество и цена** " является свертываемым и содержит количество за **единицу и** **цену за единицу**. Поле **"Дополнительные** сведения" свертывается и сворачивается при открытии карточки.

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="Снимок экрана: тип данных значения сущности с окном макета карточки. На карточке отображается заголовок карточки и разделы.":::

## <a name="card-data-attribution"></a>Атрибуция данных карточки

Карточки значений сущностей могут отображать атрибуты данных, чтобы предоставить поставщику информацию в карточке сущности. Свойство значения сущности [`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member) использует объект [`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes) , который определяет `description`и `logoSourceAddress`значения `logoTargetAddress` .

Свойство поставщика данных отображает изображение в левом нижнем углу карточки сущности. Используется для `logoSourceAddress` указания исходного URL-адреса изображения. Значение `logoTargetAddress` определяет назначение URL-адреса, если выбрано изображение логотипа. Значение `description` отображается в виде подсказки при наведении указателя мыши на логотип. Значение `description` также отображается как `logoSourceAddress` резервный текст, если объект не определен или исходный адрес изображения разбит.

В следующем фрагменте кода JSON `provider` показано значение сущности, которое использует свойство для указания атрибута поставщика данных для сущности.

> [!NOTE]
> Чтобы узнать, как использовать этот JSON в полный пример кода, посетите репозиторий [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-attribution.yaml) .

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when selected.
    }
};
```

На следующем снимке экрана показана карточка значения сущности, которая использует предыдущий фрагмент кода. Снимок экрана: атрибут поставщика данных в левом нижнем углу. В этом экземпляре поставщиком данных является корпорация Майкрософт, и отображается логотип Майкрософт.

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="Снимок экрана: тип данных значения сущности с окном макета карточки. На карточке в левом нижнем углу отображается атрибут поставщика данных.":::

## <a name="see-also"></a>См. также

- [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
- [Основные понятия, связанные с типами данных Excel](excel-data-types-concepts.md)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)