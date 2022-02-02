---
title: Предоставление надстройке разрешений администратора
description: Узнайте, как предоставить согласие администратора на надстройки
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2c3a82db390ed28c1eb8194a78f2c9fa787aeede
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320132"
---
# <a name="grant-administrator-consent-to-the-add-in"></a>Предоставление надстройке разрешений администратора

> [!NOTE]
> Эта процедура необходима только при разработке надстройки. Когда надстройка вашей продукции развернута в AppSource или Центр администрирования Microsoft 365, пользователи будут доверять ей по отдельности или администратор даст согласие для организации при установке.

Выполнить эту *процедуру после* регистрации [надстройки](../develop/register-sso-add-in-aad-v2.md).

1. Просмотрите [страницу регистрации приложений на портале Azure,](https://go.microsoft.com/fwlink/?linkid=2083908) чтобы просмотреть регистрацию приложения.

1. Вопишитесь с ***учетными*** данными администратора в Microsoft 365 аренды. Пример: MyName@contoso.onmicrosoft.com.

1. Выберите приложение с именем **$ADD-IN-NAME$**.

1. На странице **$ADD-IN-NAME$** выберите разрешения **API**, а затем в разделе Настройка разрешений выберите согласие  администратора Grant для **[имя клиента]**. Выберите **Да** для подтверждения, которое появится.

> [!NOTE]
> Рекомендуется использовать эту процедуру в качестве наилучшей практики, если вы используете Microsoft 365 [учетную запись разработчика](https://developer.microsoft.com/microsoft-365/dev-program). Однако, если вы предпочитаете, можно перезагрузить надстройку SSO в стадии разработки и подсказывать пользователю форму согласия. Дополнительные сведения см. в Windows [Sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) on [Office в Интернете](../testing/sideload-office-add-ins-for-testing.md).
