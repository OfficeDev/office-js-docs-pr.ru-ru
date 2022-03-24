---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на Office API JavaScript и определения типов в надстройки.
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 514959c7aa703172c61bcf061a9c1f047858caa4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743686"
---
# <a name="referencing-the-office-javascript-api-library"></a>Ссылки на библиотеку API JavaScript для Office 

Библиотека [API Office JavaScript предоставляет API](../reference/javascript-api-for-office.md), которые ваша надстройка может использовать для взаимодействия с Office приложением. Самый простой способ ссылки на библиотеку — использовать сеть доставки контента (CDN) `<script>` `<head>` путем добавления следующего тега в разделе вашей HTML-страницы.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

При первом загрузке надстройки загружаются и кэшируются Office API JavaScript, чтобы убедиться, что она использует самые последние реализации Office.js и связанных с ними файлов для указанной версии.

> [!IMPORTANT]
> Вы должны ссылаться на Office API `<head>` JavaScript из раздела страницы, чтобы убедиться, что API полностью инициализирован до любых элементов тела.

## <a name="api-versioning-and-backward-compatibility"></a>Версия API и обратная совместимость

В предыдущем фрагменте HTML `/1/` `office.js` перед URL-адресом CDN указывается последний дополнительный выпуск в версии 1 Office.js. Так как Office API JavaScript поддерживает обратную совместимость, последний выпуск будет по-прежнему поддерживать членов API, которые были представлены ранее в версии 1. Если необходимо обновить существующий проект, см. в статью [Обновление версии Office API JavaScript и файлы схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

Помимо ссылки на API Office JavaScript, как описано выше, вы также можете включить IntelliSense для проекта надстройки TypeScript с помощью определений типа из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого запустите следующую команду в системном запросе с поддержкой узла (или в окне баш git) из корневой папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>Предварительные API

Новые API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и необходимости отзыва пользователей.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
