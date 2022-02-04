---
title: Get and set categories
description: Управление категориями в почтовом ящике и элементе
ms.date: 01/14/2020
ms.localizationpriority: medium
---

# <a name="get-and-set-categories"></a>Get and set categories

В Outlook пользователь может применять категории к сообщениям и встречам в качестве средства организации данных почтовых ящиков. Пользователь определяет список категорий с цветным кодом для своего почтового ящика и может применить одну или несколько из этих категорий к любому элементу сообщения или встречи. Каждая [категория](/javascript/api/outlook/office.categorydetails) в мастер-списке представлена именем и цветом [,](/javascript/api/outlook/office.mailboxenums.categorycolor) указанными пользователем. Вы можете использовать API Office JavaScript для управления списком категорий в почтовом ящике и категориями, примененными к элементу.

> [!NOTE]
> Поддержка этой функции была представлена в наборе требований 1.8. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="manage-categories-in-the-master-list"></a>Управление категориями в списке master

Для применения к сообщению или встрече доступны только категории в списке магистра в почтовом ящике. API можно использовать для добавления, получения и удаления категорий магистра.

> [!IMPORTANT]
> Чтобы надстройка управляет мастер-списком категорий, `Permissions` необходимо установить узел в манифесте `ReadWriteMailbox`.

### <a name="add-master-categories"></a>Добавление категорий master

В следующем примере показано, как добавить категорию с именем "Срочно!". в мастер-список, [позвонив addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) на [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
var masterCategoriesToAdd = [
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

### <a name="get-master-categories"></a>Get master categories

В следующем примере показано, как получить список категорий, позвонив [в getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) на [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>Удаление категорий master

В следующем примере показано, как удалить категорию с именем "Срочно!". из мастер-списка по [вызову removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) на [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>Управление категориями в сообщении или встрече

API можно использовать для добавления, получения и удаления категорий для сообщения или элемента встречи.

> [!IMPORTANT]
> Для применения к сообщению или встрече доступны только категории в списке магистра в почтовом ящике. Дополнительные сведения см. в более ранней статье [Управление](#manage-categories-in-the-master-list) категориями в списке master.
>
> В Outlook в Интернете вы не можете использовать API для управления категориями в сообщении в режиме Чтения.

### <a name="add-categories-to-an-item"></a>Добавление категорий к элементу

В следующем примере показано, как применять категорию с именем "Срочно!" к текущему элементу, [позвонив в addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)).`item.categories`

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a>Получить категории элемента

В следующем примере показано, как получить категории, примененные к текущему элементу, позвонив [в getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)).`item.categories`

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>Удаление категорий из элемента

В следующем примере показано, как удалить категорию с именем "Срочно!". из текущего элемента путем вызова [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) на `item.categories`.

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a>См. также

- [Outlook разрешений](understanding-outlook-add-in-permissions.md)
- [Элемент Permissions в манифесте](../reference/manifest/permissions.md)
