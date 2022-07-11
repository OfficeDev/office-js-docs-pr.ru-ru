---
title: Получение и задание категорий
description: Управление категориями в почтовом ящике и элементе.
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: d31cb8da4cdaf4a88141a1eac927748b1399e0d9
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712827"
---
# <a name="get-and-set-categories"></a>Получение и задание категорий

В Outlook пользователь может применять категории к сообщениям и встречам в качестве средства организации данных почтового ящика. Пользователь определяет основной список цветных категорий для своего почтового ящика, а затем может применить одну или несколько из этих категорий к любому сообщению или элементу встречи. Каждая [категория в](/javascript/api/outlook/office.categorydetails) главном списке представлена именем и [цветом](/javascript/api/outlook/office.mailboxenums.categorycolor) , указанными пользователем. API JavaScript для Office можно использовать для управления главным списком категорий в почтовом ящике и категориями, примененными к элементу.

> [!NOTE]
> Поддержка этой функции реализована в наборе обязательных элементов 1.8. См [клиенты и платформы](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="manage-categories-in-the-master-list"></a>Управление категориями в главном списке

К сообщению или встрече можно применить только категории в главном списке почтового ящика. С помощью API можно добавлять, получать и удалять главные категории.

> [!IMPORTANT]
> Чтобы надстройка управляет главным списком категорий, `Permissions` необходимо задать для узла в манифесте значение `ReadWriteMailbox`.

### <a name="add-master-categories"></a>Добавление главных категорий

В следующем примере показано, как добавить категорию с именем "Срочно!" в главный список путем вызова [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) в [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a>Получение основных категорий

В следующем примере показано, как получить список категорий путем вызова [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) в [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>Удаление главных категорий

В следующем примере показано, как удалить категорию с именем "Срочно!" из главного списка путем вызова [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) в [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>Управление категориями в сообщении или встрече

С помощью API можно добавлять, получать и удалять категории для сообщения или элемента встречи.

> [!IMPORTANT]
> К сообщению или встрече можно применить только категории в главном списке почтового ящика. Дополнительные сведения см. в предыдущем разделе "Управление категориями [" в](#manage-categories-in-the-master-list) главном списке.
>
> В Outlook в Интернете вы не можете использовать API для управления категориями сообщения в режиме чтения.

### <a name="add-categories-to-an-item"></a>Добавление категорий к элементу

В следующем примере показано, как применить категорию с именем "Срочно!" для текущего элемента путем вызова [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) для `item.categories`.

```js
const categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a>Получение категорий элемента

В следующем примере показано, как получить категории, примененные к текущему элементу, вызвав [метод getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1))`item.categories`.

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>Удаление категорий из элемента

В следующем примере показано, как удалить категорию с именем "Срочно!" из текущего элемента путем вызова [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) для `item.categories`.

```js
const categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a>См. также

- [Разрешения Outlook](understanding-outlook-add-in-permissions.md)
- [Элемент Permissions в манифесте](/javascript/api/manifest/permissions)
