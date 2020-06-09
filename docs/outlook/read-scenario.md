---
title: Создание надстроек Outlook для форм чтения
description: Надстройки чтения — это надстройки Outlook, которые активируются в области чтения или с помощью инспектора чтения в Outlook.
ms.date: 04/12/2018
localization_priority: Priority
ms.openlocfilehash: 815234ed046b4c00b91f5acd6cd2c4dcd226dba2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605306"
---
# <a name="create-outlook-add-ins-for-read-forms"></a><span data-ttu-id="4161b-103">Создание надстроек Outlook для форм чтения</span><span class="sxs-lookup"><span data-stu-id="4161b-103">Create Outlook add-ins for read forms</span></span>

<span data-ttu-id="4161b-p101">Надстройки чтения — это надстройки Outlook, которые активируются в области чтения или инспекторе в Outlook. В отличие от надстроек создания, которые активируются, когда пользователь создает сообщение или встречу, надстройки чтения доступны, когда пользователь:</span><span class="sxs-lookup"><span data-stu-id="4161b-p101">Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:</span></span> 

- <span data-ttu-id="4161b-106">Просматривает электронное сообщение, приглашение на собрание, ответ на приглашение или уведомление об отмене собрания.</span><span class="sxs-lookup"><span data-stu-id="4161b-106">View an email message, meeting request, meeting response, or meeting cancellation.</span></span>

   > [!NOTE]
   > <span data-ttu-id="4161b-107">Outlook не активирует надстройки в форме чтения для некоторых типов сообщений, в том числе элементов, являющихся вложениями в других сообщениях, элементов в папке черновиков, а также зашифрованных или защищенных другим способом элементов.</span><span class="sxs-lookup"><span data-stu-id="4161b-107">Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.</span></span>
    
- <span data-ttu-id="4161b-108">Просматривает данные о собрании в роли участника.</span><span class="sxs-lookup"><span data-stu-id="4161b-108">View a meeting item in which the user is an attendee.</span></span>
    
- <span data-ttu-id="4161b-109">просматривает собрание в роли организатора (только выпуск RTM Outlook 2013 и Exchange 2013).</span><span class="sxs-lookup"><span data-stu-id="4161b-109">View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="4161b-p102">Начиная с выпуска Office 2013 с пакетом обновления SP1, если пользователь просматривает данные об организованном им собрании, доступны только надстройки создания. Надстройки чтения больше не доступны в этом сценарии.</span><span class="sxs-lookup"><span data-stu-id="4161b-p102">Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.</span></span>


<span data-ttu-id="4161b-p103">В каждом из этих сценариев чтения Outlook активирует надстройки при выполнении заданных условий, а пользователи могут открыть активные надстройки в поле надстроек инспектора или области чтения. На приведенном ниже рисунке показана надстройка **Карты Bing**, которая активируется и открывается, когда пользователь читает сообщение, содержащее географический адрес.</span><span class="sxs-lookup"><span data-stu-id="4161b-p103">In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.</span></span>


<span data-ttu-id="4161b-114">**Область надстройки с надстройкой "Карты Bing" для выбранного сообщения Outlook, содержащего адрес**</span><span class="sxs-lookup"><span data-stu-id="4161b-114">**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**</span></span>

![Почтовое приложение "Карты Bing" в Outlook](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a><span data-ttu-id="4161b-116">Типы надстроек, доступные в режиме чтения</span><span class="sxs-lookup"><span data-stu-id="4161b-116">Types of add-ins available in read mode</span></span>

<span data-ttu-id="4161b-117">Надстройки чтения могут быть любым сочетанием следующих типов:</span><span class="sxs-lookup"><span data-stu-id="4161b-117">Read add-ins can be any combination of the following types.</span></span>

- [<span data-ttu-id="4161b-118">Команды надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="4161b-118">Add-in commands for Outlook</span></span>](add-in-commands-for-outlook.md)   
- [<span data-ttu-id="4161b-119">Контекстные надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4161b-119">Contextual Outlook add-ins</span></span>](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a><span data-ttu-id="4161b-120">Функции API, доступные надстройкам чтения</span><span class="sxs-lookup"><span data-stu-id="4161b-120">API features available to read add-ins</span></span>

- <span data-ttu-id="4161b-121">Активация надстроек в формах чтения: см. таблицу 1 в разделе [Указание правил активации в манифесте](activation-rules.md#specify-activation-rules-in-a-manifest).</span><span class="sxs-lookup"><span data-stu-id="4161b-121">For activating add-ins in read forms: see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).</span></span>    
- [<span data-ttu-id="4161b-122">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4161b-122">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [<span data-ttu-id="4161b-123">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="4161b-123">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="4161b-124">Извлечение строк сущностей из элемента Outlook</span><span class="sxs-lookup"><span data-stu-id="4161b-124">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)   
- [<span data-ttu-id="4161b-125">Получение вложений элемента Outlook с сервера</span><span class="sxs-lookup"><span data-stu-id="4161b-125">Get attachments of an Outlook item from the server</span></span>](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a><span data-ttu-id="4161b-126">См. также</span><span class="sxs-lookup"><span data-stu-id="4161b-126">See also</span></span>

- [<span data-ttu-id="4161b-127">Написание вашей первой надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="4161b-127">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
