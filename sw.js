<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>To-Do List PWA</title>
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Excel processing -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <!-- Link to manifest for PWA -->
  <link rel="manifest" href="manifest.json">
  <!-- Theme color for PWA -->
  <meta name="theme-color" content="#3B82F6">
</head>
<body class="bg-gray-100 font-sans">
  <div class="container mx-auto p-4 max-w-lg">
    <h1 class="text-3xl font-bold text-center mb-6 text-gray-800">To-Do List</h1>
    
    <!-- Input Form -->
    <div class="flex flex-col space-y-2 mb-4">
      <input 
        id="taskInput" 
        type="text" 
        placeholder="Enter a task" 
        class="p-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
      >
      <select id="categorySelect" class="p-2 border border-gray-300 rounded-lg">
        <option value="Work">Work</option>
        <option value="Personal">Personal</option>
        <option value="Other">Other</option>
      </select>
      <label class="flex items-center space-x-2">
        <input 
          id="addDueDateCheckbox" 
          type="checkbox" 
          class="h-5 w-5 text-blue-500 focus:ring-blue-500 border-gray-300 rounded"
          onchange="toggleDueDateInput()"
        >
        <span class="text-gray-700">Add Due Date</span>
      </label>
      <input 
        id="dueDateInput" 
        type="datetime-local" 
        class="hidden p-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
      >
      <button 
        onclick="addTask()" 
        class="bg-blue-500 text-white p-2 rounded-lg hover:bg-blue-600"
      >
        Add Task
      </button>
    </div>

    <!-- Task List -->
    <ul id="taskList" class="space-y-2"></ul>
  </div>

  <script type="text/javascript">
    // Excel processing script (unchanged)
    var gk_isXlsx = false;
    var gk_xlsxFileLookup = {};
    var gk_fileData = {};
    function filledCell(cell) {
      return cell !== '' && cell != null;
    }
    function loadFileData(filename) {
      if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
        try {
          var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
          var firstSheetName = workbook.SheetNames[0];
          var worksheet = workbook.Sheets[firstSheetName];
          var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
          var filteredData = jsonData.filter(row => row.some(filledCell));
          var headerRowIndex = filteredData.findIndex((row, index) =>
            row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
          );
          if (headerRowIndex === -1 || headerRowIndex > 25) {
            headerRowIndex = 0;
          }
          var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex));
          csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
          return csv;
        } catch (e) {
          console.error(e);
          return "";
        }
      }
      return gk_fileData[filename] || "";
    }
  </script>

  <script>
    // Load tasks from localStorage on page load
    let tasks = JSON.parse(localStorage.getItem('tasks')) || [];

    // Format date for display
    function formatDate(dateString) {
      if (!dateString) return 'No due date';
      const date = new Date(dateString);
      return date.toLocaleString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    }

    // Check if a due date is overdue
    function isOverdue(dueDate) {
      if (!dueDate) return false;
      return new Date(dueDate).getTime() < new Date().getTime();
    }

    // Toggle due date input visibility
    function toggleDueDateInput() {
      const checkbox = document.getElementById('addDueDateCheckbox');
      const dueDateInput = document.getElementById('dueDateInput');
      dueDateInput.classList.toggle('hidden', !checkbox.checked);
      if (!checkbox.checked) {
        dueDateInput.value = ''; // Clear input when hidden
      }
    }

    // Render tasks
    function renderTasks() {
      const taskList = document.getElementById('taskList');
      taskList.innerHTML = '';
      tasks.forEach((task, index) => {
        const li = document.createElement('li');
        li.className = 'flex flex-col p-2 bg-white border border-gray-200 rounded-lg shadow-sm';
        
        // Task content
        const taskContent = document.createElement('div');
        taskContent.className = 'flex justify-between items-center';
        const taskSpan = document.createElement('span');
        taskSpan.textContent = `${task.text} [${task.category}]`;
        taskSpan.className = 'cursor-pointer';
        taskSpan.style.textDecoration = task.completed ? 'line-through' : 'none';
        taskSpan.onclick = () => {
          tasks[index].completed = !tasks[index].completed;
          saveTasks();
          renderTasks();
        };
        
        // Delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Delete';
        deleteBtn.className = 'bg-red-500 text-white px-2 py-1 rounded hover:bg-red-600';
        deleteBtn.onclick = () => {
          tasks.splice(index, 1);
          saveTasks();
          renderTasks();
        };

        taskContent.appendChild(taskSpan);
        taskContent.appendChild(deleteBtn);

        // Due date display
        const dueDateDiv = document.createElement('div');
        dueDateDiv.textContent = `Due: ${formatDate(task.dueDate)}`;
        dueDateDiv.className = `text-sm mt-1 ${isOverdue(task.dueDate) ? 'text-red-600' : 'text-gray-600'}`;

        // Reminder button
        const reminderBtn = document.createElement('button');
        reminderBtn.textContent = 'Set Reminder';
        reminderBtn.className = 'mt-2 bg-green-500 text-white px-2 py-1 rounded hover:bg-green-600';
        reminderBtn.onclick = () => setReminder(task.text, task.dueDate);

        li.appendChild(taskContent);
        li.appendChild(dueDateDiv);
        li.appendChild(reminderBtn);
        taskList.appendChild(li);
      });
    }

    // Save tasks to localStorage
    function saveTasks() {
      localStorage.setItem('tasks', JSON.stringify(tasks));
    }

    // Add a task
    function addTask() {
      const taskInput = document.getElementById('taskInput');
      const categorySelect = document.getElementById('categorySelect');
      const dueDateCheckbox = document.getElementById('addDueDateCheckbox');
      const dueDateInput = document.getElementById('dueDateInput');
      const taskText = taskInput.value.trim();
      const category = categorySelect.value;
      const dueDate = dueDateCheckbox.checked ? dueDateInput.value : null;

      if (taskText === '') {
        alert('Please enter a task!');
        return;
      }

      if (dueDateCheckbox.checked && !dueDate) {
        alert('Please select a due date or uncheck the "Add Due Date" option!');
        return;
      }

      if (dueDate) {
        const dueDateTime = new Date(dueDate).getTime();
        const now = new Date().getTime();
        const oneYearFromNow = now + 365 * 24 * 60 * 60 * 1000; // 1 year in milliseconds

        if (dueDateTime < now) {
          alert('Due date cannot be in the past!');
          return;
        }

        if (dueDateTime > oneYearFromNow) {
          alert('Due date cannot be more than 1 year in the future!');
          return;
        }
      }

      tasks.push({
        text: taskText,
        category: category,
        dueDate: dueDate, // Store null if no due date
        completed: false
      });
      saveTasks();
      renderTasks();
      taskInput.value = '';
      dueDateInput.value = '';
      dueDateCheckbox.checked = false;
      toggleDueDateInput(); // Hide due date input after adding
    }

    // Set a reminder notification
    function setReminder(taskText, dueDate) {
      if (!("Notification" in window)) {
        alert("This browser does not support notifications.");
        return;
      }

      // Require a due date for reminders
      if (!dueDate) {
        alert("A due date is required to set a reminder.");
        return;
      }

      const dueDateTime = new Date(dueDate).getTime();
      const now = new Date().getTime();
      const delay = dueDateTime - now;

      if (delay <= 0) {
        alert("The due date is in the past. Please set a future due date for the reminder.");
        return;
      }

      if (Notification.permission === "granted") {
        setTimeout(() => {
          new Notification("To-Do Reminder", {
            body: `Time to work on: ${taskText}\nDue: ${formatDate(dueDate)}`,
            icon: 'icon-192x192.png'
          });
        }, delay);
        alert(`Reminder set for ${formatDate(dueDate)}`);
      } else if (Notification.permission !== "denied") {
        Notification.requestPermission().then(permission => {
          if (permission === "granted") {
            setTimeout(() => {
              new Notification("To-Do Reminder", {
                body: `Time to work on: ${taskText}\nDue: ${formatDate(dueDate)}`,
                icon: 'icon-192x192.png'
              });
            }, delay);
            alert(`Reminder set for ${formatDate(dueDate)}`);
          }
        });
      }
    }

    // Allow adding task with Enter key
    document.getElementById('taskInput').addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        addTask();
      }
    });

    // Register service worker for PWA
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('/sw.js')
          .then(reg => console.log('Service Worker registered'))
          .catch(err => console.log('Service Worker registration failed:', err));
      });
    }

    // Initial render
    renderTasks();
  </script>
</body>
</html>
