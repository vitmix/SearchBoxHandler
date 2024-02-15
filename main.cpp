#include "UIAutomationStuff.h"
#include <atomic>
#include <chrono>

static const std::string stopWord("quit");

void HandleUserInput(std::atomic_bool& run)
{
    std::string input;

    while (run.load())
    {
        std::cin >> input;
        if (input == stopWord)
            run.store(false);
    }
}

int main()
{
    using namespace std::chrono_literals;

    // SyncBlock is used for synchronization between browser window opened event
    // handler and main thread launching event handlers for url manipulation
    utils::SyncBlock sBlock;

    uia::UIManager uiManager;
    
    std::wcout << "Initializing..." << std::endl;
    
    if (!uiManager.init(sBlock))
    {
        std::wcout << "Failed to init UI Manager" << std::endl;
        return 1;
    }

    std::wcout << "Print \"quit\" to stop url manipulator" << std::endl;

    // Launch separate thread to handle user input
    std::atomic_bool run(true);
    std::thread userInputThread(HandleUserInput, std::ref(run));

    while (run.load())
    {
        // wait for the new browser window opened for 1 sec and try to add 
        // other Edit Control + event handler
        {
            std::unique_lock lock(sBlock.mx);
            const auto status = sBlock.cv.wait_for(lock, 1000ms, [&] { return sBlock.processed; });
            if (!status)
            {
                // try to sleep for another 1 sec and then continue
                std::this_thread::sleep_for(1000ms);
                continue;
            }
        }
        uiManager.tryAddNewHandler();

        // need to reset boolean to be able to wait for the new window
        {
            std::lock_guard lock(sBlock.mx);
            sBlock.processed = false;
        }
    }

    run.store(false);
    userInputThread.join();

    std::wcout << "Finished processing." << std::endl;

    return 0;
}
