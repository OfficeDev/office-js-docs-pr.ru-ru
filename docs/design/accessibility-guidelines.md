---
title: Рекомендации по специальным возможностям в надстройках Office
description: ''
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: c40ca0c3c1fad238d605e5f3f67b58a0272ff83a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448135"
---
# <a name="accessibility-guidelines"></a><span data-ttu-id="2a06a-102">Рекомендации по специальным возможностям</span><span class="sxs-lookup"><span data-stu-id="2a06a-102">Accessibility guidelines</span></span>

<span data-ttu-id="2a06a-p101">При проектировании и разработке надстроек Office вам следует обеспечить возможность успешного использования этих решений для всех потенциальных пользователей и клиентов. Следуйте приведенным ниже рекомендациям, чтобы обеспечить доступность вашего решения для максимально широкой аудитории.</span><span class="sxs-lookup"><span data-stu-id="2a06a-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="2a06a-105">Использование нескольких способов ввода</span><span class="sxs-lookup"><span data-stu-id="2a06a-105">Design for multiple input methods</span></span>

- <span data-ttu-id="2a06a-p102">Убедитесь, что пользователи могут выполнять операции, используя только клавиатуру. У пользователей должна быть возможность переходить ко всем активным элементам на странице с помощью сочетания клавиши TAB и клавиш со стрелками.</span><span class="sxs-lookup"><span data-stu-id="2a06a-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="2a06a-108">Если элементы управления используются на мобильном устройстве, то при их касании должны воспроизводиться соответствующие звуковые сигналы.</span><span class="sxs-lookup"><span data-stu-id="2a06a-108">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="2a06a-109">Предоставьте полезные подписи для всех интерактивных элементов управления.</span><span class="sxs-lookup"><span data-stu-id="2a06a-109">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="2a06a-110">Простота использования надстройки</span><span class="sxs-lookup"><span data-stu-id="2a06a-110">Make your add-in easy to use</span></span>

- <span data-ttu-id="2a06a-111">При создании элементов пользовательского интерфейса не полагайтесь только на один атрибут, например цвет, размер, форму, расположение, ориентацию или звук.</span><span class="sxs-lookup"><span data-stu-id="2a06a-111">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="2a06a-112">Избегайте непредвиденных изменений контекста без участия пользователя, например перемещения фокуса к другому элементу пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="2a06a-112">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="2a06a-113">Предоставляйте возможность проверки, подтверждения или аннулирования всех действий, к чему-либо обязывающих.</span><span class="sxs-lookup"><span data-stu-id="2a06a-113">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="2a06a-114">Предоставляйте возможность приостановки или остановки воспроизведения мультимедиа, например аудио и видео.</span><span class="sxs-lookup"><span data-stu-id="2a06a-114">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="2a06a-115">Не ограничивайте по времени действия пользователей.</span><span class="sxs-lookup"><span data-stu-id="2a06a-115">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="2a06a-116">Наглядность при работе с надстройкой</span><span class="sxs-lookup"><span data-stu-id="2a06a-116">Make your add-in easy to see</span></span>

- <span data-ttu-id="2a06a-117">Избегайте непредвиденных изменений цвета.</span><span class="sxs-lookup"><span data-stu-id="2a06a-117">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="2a06a-p103">Предоставляйте понятную и своевременную информацию для описания элементов пользовательского интерфейса, названий и заголовков, входных данных и ошибок. Убедитесь, что имена элементов управления адекватно описывают их предназначение.</span><span class="sxs-lookup"><span data-stu-id="2a06a-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="2a06a-120">Следуйте [стандартным рекомендациям](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) по использованию цветового контраста.</span><span class="sxs-lookup"><span data-stu-id="2a06a-120">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="2a06a-121">Учет специальных возможностей</span><span class="sxs-lookup"><span data-stu-id="2a06a-121">Account for assistive technologies</span></span>

- <span data-ttu-id="2a06a-122">Не используйте функции, которые не сочетаются со специальными возможностями, в том числе визуальными, звуковыми или другими типами взаимодействия.</span><span class="sxs-lookup"><span data-stu-id="2a06a-122">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="2a06a-p104">Не используйте текст в формате изображений. Средства чтения с экрана не могут считывать текст внутри изображений.</span><span class="sxs-lookup"><span data-stu-id="2a06a-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="2a06a-125">Предоставляйте пользователям возможность настраивать или отключать все источники звука.</span><span class="sxs-lookup"><span data-stu-id="2a06a-125">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="2a06a-126">Предоставляйте пользователям возможность включения титров или звукового описания вместе с источниками звука.</span><span class="sxs-lookup"><span data-stu-id="2a06a-126">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="2a06a-127">Для предупреждения пользователей помимо звуковых сигналов должны быть доступны и другие варианты, например визуальные подсказки или вибрация.</span><span class="sxs-lookup"><span data-stu-id="2a06a-127">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="2a06a-128">См. также</span><span class="sxs-lookup"><span data-stu-id="2a06a-128">See also</span></span>

- [<span data-ttu-id="2a06a-129">WCAG (Web Content Accessibility Guidelines) 2.0</span><span class="sxs-lookup"><span data-stu-id="2a06a-129">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- <span data-ttu-id="2a06a-130">[Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)](https://www.w3.org/TR/wcag2ict/) (Руководство по использованию WCAG 2.0 с информационными и коммуникационными технологиями, не связанными с Интернетом)</span><span class="sxs-lookup"><span data-stu-id="2a06a-130">[Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)](https://www.w3.org/TR/wcag2ict/)</span></span>
- <span data-ttu-id="2a06a-131">[European Standard on accessibility requirements for Information and Communication Technologies (ICT)](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) (Европейский стандарт относительно требований к специальным возможностям для информационных и коммуникационных технологий)</span><span class="sxs-lookup"><span data-stu-id="2a06a-131">[European Standard on accessibility requirements for Information and Communication Technologies (ICT)](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf)</span></span> 
