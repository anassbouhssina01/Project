/**
 * @file Main application logic for the Employee Invitation Tool.
 * Handles data loading, UI interactions, search, list management,
 * and Word document generation.
 * @author Jules
 */

document.addEventListener('DOMContentLoaded', () => {
    console.log("app.js loaded");

    // --- DOM Element Selectors ---
    const jsonUpload = document.getElementById('json-upload');
    const toggleJsonUploadBtn = document.getElementById('toggle-json-upload');
    const jsonUploadContainer = document.getElementById('json-upload-container');

    const excelUpload = document.getElementById('excel-upload');
    const toggleExcelUploadBtn = document.getElementById('toggle-excel-upload');
    const excelUploadContainer = document.getElementById('excel-upload-container');

    const searchInput = document.getElementById('search-input');
    const searchCityInput = document.getElementById('search-city');
    const searchResults = document.getElementById('search-results');
    
    const invitedListDiv = document.getElementById('invited-list');
    const wordTemplateUpload = document.getElementById('word-template-upload');
    const generateWordBtn = document.getElementById('generate-word');
    const clearAllBtn = document.getElementById('clear-all-btn');

    // --- Application State Variables ---
    let employeeData = []; // Holds the master list of all employees
    let invitedList = [];  // Holds the list of employees selected for invitation
    let wordTemplate = null; // Holds the binary content of the uploaded .docx template

    // --- UI Visibility Toggles ---
    toggleJsonUploadBtn.addEventListener('click', () => {
        jsonUploadContainer.classList.toggle('hidden');
    });

    toggleExcelUploadBtn.addEventListener('click', () => {
        excelUploadContainer.classList.toggle('hidden');
    });

    // --- Data Persistence ---

    /**
     * Loads the master employee data from localStorage.
     */
    function loadEmployeeData() {
        const data = localStorage.getItem('employeeData');
        employeeData = data ? JSON.parse(data) : [];
    }

    /**
     * Loads the list of invited employees from localStorage and re-renders the list.
     */
    function loadInvitedList() {
        const data = localStorage.getItem('invitedListData');
        invitedList = data ? JSON.parse(data) : [];
        renderInvitedList();
    }

    /**
     * Saves the current list of invited employees to localStorage.
     */
    function saveInvitedList() {
        localStorage.setItem('invitedListData', JSON.stringify(invitedList));
    }

    // --- JSON Data Upload ---

    /**
     * Handles the master employee data upload from a JSON file.
     */
    jsonUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (data.employees && Array.isArray(data.employees)) {
                    localStorage.setItem('employeeData', JSON.stringify(data.employees));
                    loadEmployeeData();
                    alert('تم تحميل بيانات الموظفين بنجاح!');
                    jsonUploadContainer.classList.add('hidden');
                } else {
                    alert('ملف JSON غير صالح. يجب أن يحتوي على مصفوفة باسم "employees".');
                }
            } catch (error) {
                console.error('Error parsing JSON:', error);
                alert('حدث خطأ أثناء قراءة الملف. الرجاء التأكد من أن الملف بصيغة JSON صحيحة.');
            }
        };
        reader.readAsText(file);
    });

    // --- Search Functionality ---

    /**
     * Filters the master employee list based on search terms and displays the results.
     */
    function performSearch() {
        const searchTerm = searchInput.value.toLowerCase().trim();
        const cityTerm = searchCityInput.value.toLowerCase().trim();
        searchResults.innerHTML = '';

        if (searchTerm.length === 0 && cityTerm.length === 0) {
            return;
        }

        const filteredEmployees = employeeData.filter(emp => {
            const empName = (emp.fullName || '').toLowerCase();
            const empId = (emp.employeeId || '').toString();
            const empCity = (emp.city || '').toLowerCase();
            const postResponsibility = (emp.postResponsibility || '').trim();

            const nameMatch = searchTerm ? (empName.includes(searchTerm) || empId.includes(searchTerm)) : true;
            const cityMatch = cityTerm ? empCity.includes(cityTerm) : true;

            // Special filter for "responsible" people if searching by city.
            const responsibleMatch = cityTerm 
                ? (postResponsibility.includes("رئيس") || postResponsibility.includes("رئيسة") || postResponsibility.includes("وكيل") || postResponsibility.includes("وكيلة"))
                : true;

            return nameMatch && cityMatch && responsibleMatch;
        });

        displaySearchResults(filteredEmployees);
    }

    searchInput.addEventListener('keyup', performSearch);
    searchCityInput.addEventListener('keyup', performSearch);

    /**
     * Renders the search results in the UI.
     * @param {Array<Object>} employees - The array of employees to display.
     */
    function displaySearchResults(employees) {
        searchResults.innerHTML = '';
        employees.forEach(emp => {
            const empDiv = document.createElement('div');
            empDiv.className = 'employee-item';
            empDiv.innerHTML = `
                <div class="employee-info">
                    <span>${emp.fullName} (الرقم: ${emp.employeeId})</span>
                    <span class="details">${emp.workLocation} | ${emp.city} | ${emp.division}</span>
                </div>
            `;
            empDiv.dataset.employeeId = emp.employeeId;
            
            empDiv.addEventListener('click', () => {
                addToInvitedList(emp.employeeId);
            });

            searchResults.appendChild(empDiv);
        });
    }

    // --- Invited List Management ---

    /**
     * Renders the list of invited employees, grouped by location/division/city.
     * Highlights "responsible" employees with a special class.
     */
    function renderInvitedList() {
        invitedListDiv.innerHTML = '';

        if (invitedList.length === 0) {
            return;
        }

        // 1. Group employees by a composite key for display.
        const groupedInvitedList = invitedList.reduce((acc, emp) => {
            const key = `${emp.workLocation || 'N/A'}_${emp.division || 'N/A'}_${emp.city || 'N/A'}`;
            if (!acc[key]) {
                acc[key] = {
                    workLocation: emp.workLocation,
                    division: emp.division,
                    city: emp.city,
                    employees: []
                };
            }
            acc[key].employees.push(emp);
            return acc;
        }, {});

        // 2. Render each group with a header.
        for (const key in groupedInvitedList) {
            const group = groupedInvitedList[key];
            
            const groupHeader = document.createElement('h3');
            groupHeader.className = 'group-header';
            groupHeader.textContent = `${group.workLocation || ''} - ${group.division || ''} - ${group.city || ''}`;
            invitedListDiv.appendChild(groupHeader);

            // Render each employee within the group.
            group.employees.forEach(emp => {
                const item = document.createElement('div');
                const responsibleClass = isResponsible(emp) ? ' responsible' : '';
                item.className = `employee-item${responsibleClass}`;
                
                item.innerHTML = `
                    <div class="employee-info">
                        <span>${emp.jobTitle || ''} ${emp.fullName} (الرقم: ${emp.employeeId})</span>
                        <span class="details">${emp.workLocation} | ${emp.division} | ${emp.city}</span>
                    </div>
                    <button class="delete-btn" data-employee-id="${emp.employeeId}">حذف</button>
                `;
                invitedListDiv.appendChild(item);
            });
        }
    }

    /**
     * Adds an employee to the invited list by their ID.
     * @param {string|number} employeeId - The ID of the employee to add.
     */
    function addToInvitedList(employeeId) {
        const employeeToAdd = employeeData.find(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10));
        
        if (employeeToAdd && !invitedList.some(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10))) {
            invitedList.push(employeeToAdd);
            saveInvitedList();
            renderInvitedList();
            // alert(`تمت إضافة "${employeeToAdd.fullName}" إلى لائحة المدعوين.`);
            // Clear search inputs and results after adding.
            searchInput.value = '';
            searchCityInput.value = '';
            searchResults.innerHTML = '';
        } else if (!employeeToAdd) {
             alert('لم يتم العثور على الموظف.');
        } else {
            alert('هذا الموظف موجود بالفعل في القائمة.');
        }
    }

    /**
     * Removes an employee from the invited list by their ID.
     * @param {string|number} employeeId - The ID of the employee to remove.
     */
    function removeFromInvitedList(employeeId) {
        invitedList = invitedList.filter(emp => parseInt(emp.employeeId, 10) !== parseInt(employeeId, 10));
        saveInvitedList();
        renderInvitedList();
    }

    /**
     * Checks if an employee is a "responsible person" based on their job title.
     * @param {Object} employee - The employee object.
     * @returns {boolean} - True if the employee is a responsible person.
     */
    function isResponsible(employee) {
        if (!employee || !employee.postResponsibility) return false;
        const title = employee.postResponsibility.trim();
        return (title.includes("رئيس") || title.includes("رئيسة") || title.includes("وكيل") || title.includes("وكيلة"));
    }

    // Event listener for the delete buttons on individual invited employees.
    invitedListDiv.addEventListener('click', (event) => {
        if (event.target.classList.contains('delete-btn')) {
            const employeeId = event.target.dataset.employeeId;
            removeFromInvitedList(employeeId);
        }
    });

    /**
     * Event listener for the "Clear All" button. Clears the entire invited list.
     */
    clearAllBtn.addEventListener('click', () => {
        if (invitedList.length > 0 && confirm('هل أنت متأكد أنك تريد مسح لائحة المدعوين بالكامل؟')) {
            invitedList = [];
            saveInvitedList();
            renderInvitedList();
            alert('تم مسح لائحة المدعوين بنجاح.');
        }
    });

    // --- Excel Import ---

    /**
     * Handles importing a list of employee IDs from an Excel file.
     * The imported list REPLACES the current invited list after user confirmation.
     * Expects a headerless Excel file with IDs in the first column.
     */
     // --- Excel Import ---
    excelUpload.addEventListener('change', (event) => {
        if (typeof XLSX === 'undefined') {
            alert('عذرًا، حدث خطأ أثناء تحميل مكتبة معالجة ملفات Excel.');
            return;
        }

        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                // Use { header: 1 } to get an array of arrays, suitable for headerless files.
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                const employeesToAdd = [];
                rows.forEach(row => {
                    // The employee ID is in the first column of each row.
                    const employeeId = row[0];
                    if (employeeId) {
                        const foundEmployee = employeeData.find(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10));
                        
                        // Check if employee exists and is not already pending to be added.
                        const isAlreadyPending = employeesToAdd.some(emp => parseInt(emp.employeeId, 10) === parseInt(employeeId, 10));

                        if (foundEmployee && !isAlreadyPending) {
                            employeesToAdd.push(foundEmployee);
                        }
                    }
                });

                if (employeesToAdd.length > 0) {
                    const employeeNames = employeesToAdd.map(emp => emp.fullName).join('\n');
                    const confirmationMessage = `
                        سيتم مسح القائمة الحالية واستبدالها بـ ${employeesToAdd.length} موظف جديد من الملف. هل تريد المتابعة؟
                        \n\n${employeeNames}
                    `;
                    
                    if (window.confirm(confirmationMessage)) {
                        invitedList = employeesToAdd; // Replace the old list with the new one
                        saveInvitedList();
                        renderInvitedList();
                        alert(`تم استبدال لائحة المدعوين بنجاح.`);
                    } else {
                        alert('تم إلغاء عملية الاستبدال.');
                    }
                } else {
                    alert('لم يتم العثور على موظفين صالحين في ملف Excel لإنشاء قائمة جديدة.');
                }

            } catch (error) {
                console.error("Error processing Excel file:", error);
                alert("حدث خطأ أثناء معالجة ملف Excel.");
            } finally {
                excelUpload.value = '';
                excelUploadContainer.classList.add('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // --- Word Document Generation ---
    wordTemplateUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            wordTemplate = null;
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            wordTemplate = e.target.result;
            alert('تم تحميل قالب Word بنجاح.');
        };
        reader.readAsArrayBuffer(file);
    });

    // --- Word Document Generation ---

    /**
     * Handles the upload of the .docx template file.
     */
    wordTemplateUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            wordTemplate = null;
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            wordTemplate = e.target.result;
            alert('تم تحميل قالب Word بنجاح.');
        };
        reader.readAsArrayBuffer(file);
    });
    
    // Using the user-specified, smaller dictionary.
    const jobTitleGrammar = {
        'قاضي': { m: {s: 'القاضي ب', d: 'القاضيين ب', p: 'القضاة ب'}, f: {s: 'القاضية ب', d: 'القاضيتين ب', p: 'القاضيات ب'} },
        'مستشار': { m: {s: 'المستشار ب', d: 'المستشارين ب', p: 'المستشارون ب'}, f: {s: 'المستشارة ب', d: 'المستشارتين ب', p: 'المستشارات ب'} },
        'محام عام': { m: {s: 'محام عام', d: 'محاميان عامان', p: 'محامون عامون'}, f: {s: 'محامية عامة', d: 'محاميتان عامتان', p: 'محاميات عامات'} },
        'نائب الوكيل العام للملك': { m: {s: ' نائب الوكيل العام للملك لدى', d: ' نائبا الوكيل العام للملك لدى', p: 'نواب الوكيل العام للملك لدى '}, f: {s: 'نائبة الوكيل العام للملك لدى ', d: 'لنائبتا الوكيل العام للملك لدى ', p: 'نائبات الوكيل العام للملك لدى '} },
        'نائب وكيل الملك': { m: {s: 'نائب وكيل الملك لدى', d: 'نائبا وكيل الملك لدى', p: 'نواب وكيل الملك لدى '}, f: {s: 'نائبة وكيل الملك لدى ', d: 'نائبتا وكيل الملك لدى ', p: 'نائبات وكيل الملك لدى '} },
    };

    /**
     * Finds the masculine root key for a given job title from the grammar dictionary.
     * @param {string} title - The job title to find the root for.
     * @returns {string} - The masculine root key from the dictionary, or the original title if not found.
     */
    function findMasculineRoot(title) {
        if (!title) return '';
        // Check for a direct match (i.e., the title is already the masculine root)
        if (jobTitleGrammar[title]) {
            return title;
        }
        // If not, iterate through the dictionary to find a match for a feminine form.
        for (const key in jobTitleGrammar) {
            const grammar = jobTitleGrammar[key];
            if (grammar.f.s === title || grammar.f.d === title || grammar.f.p === title) {
                return key; // Return the masculine root key
            }
        }
        // Fallback for compound titles where the feminine marker is on the first word.
        // This is a more robust way to handle titles like "نائبة وكيل الملك".
        const firstWord = title.split(' ')[0];
        if (firstWord.endsWith('ة')) {
            const masculinizedTitle = title.replace(firstWord, firstWord.slice(0, -1));
            if (jobTitleGrammar[masculinizedTitle]) {
                return masculinizedTitle;
            }
        }
        return title; // Return the original title if no key is found
    }


    /**
     * Processes a group of employees to generate grammatically correct collective titles.
     * @param {Array<Object>} employees - A list of employees in a single group.
     * @returns {Object} - An object containing various processed strings for the template.
     */
    function processGroupData(employees) {
        const males = employees.filter(e => e.gender === 'السيد');
        const females = employees.filter(e => e.gender === 'السيدة');
        
        // Generate base title for males
        let maleCollectiveTitle = '';
        if (males.length > 0) {
            const title = males.length === 1 ? 'السيد' : (males.length === 2 ? 'السيدين' : 'السادة');
            const maleNames = males.map(e => e.fullName).join(' و ');
            maleCollectiveTitle = `${title} ${maleNames}`;
        }

        // Generate base title for females
        let femaleCollectiveTitle = '';
        if (females.length > 0) {
            const title = females.length === 1 ? 'السيدة' : (females.length === 2 ? 'السيدتين' : 'السيدات');
            const femaleNames = females.map(e => e.fullName).join(' و ');
            femaleCollectiveTitle = `${title} ${femaleNames}`;
        }

        // Add 'ب' prefix to the beginning of the collective title string.
        if (males.length > 0) {
            maleCollectiveTitle = `ب${maleCollectiveTitle}`;
        } else if (females.length > 0) {
            femaleCollectiveTitle = `ب${femaleCollectiveTitle}`;
        }

        // Add 'و' connector if the group is mixed.
        if (males.length > 0 && females.length > 0) {
            femaleCollectiveTitle = ` و ${femaleCollectiveTitle}`;
        }

        // Generate a context-dependent variable (e.g., "المعني", "المعنيان", "المعنيين")
        let computedVar = '';
        const totalCount = employees.length;
        if (totalCount === 1) {
            computedVar = males.length === 1 ? 'المعني' : 'المعنية';
        } else if (totalCount === 2) {
            if (males.length === 2) computedVar = 'المعنيين';
            else if (females.length === 2) computedVar = 'المعنيتين';
            else computedVar = 'المعنيين';
        } else {
             computedVar = males.length > 0 ? 'المعنيين' : 'المعنيات';
        }

        const city = employees.length > 0 ? employees[0].city : '';

        // This logic determines the correct grammatical form of the job title for the group.
        let combinedJobTitle = '';
        if (employees.length > 0) {
            const representativeTitle = (employees[0].jobTitle || '').trim();
            const root = findMasculineRoot(representativeTitle);

            if (jobTitleGrammar[root]) {
                const grammarSet = jobTitleGrammar[root];
                const groupMaleCount = males.length;
                const groupFemaleCount = females.length;
                const groupTotalCount = employees.length;

                if (groupTotalCount === 1) {
                    combinedJobTitle = groupMaleCount === 1 ? grammarSet.m.s : grammarSet.f.s;
                } else if (groupTotalCount === 2) {
                    if (groupFemaleCount === 2) {
                        combinedJobTitle = grammarSet.f.d;
                    } else {
                        combinedJobTitle = grammarSet.m.d;
                    }
                } else if (groupTotalCount > 2){
                    if (groupMaleCount > 0) {
                        combinedJobTitle = grammarSet.m.p;
                    } else {
                        combinedJobTitle = grammarSet.f.p;
                    }
                } else {
                    combinedJobTitle = root || '';
                }
            } else {
                combinedJobTitle = representativeTitle || ''; // Fallback to the original title
            }
        }
        
        // A simple string of all names, joined by 'and'.
        const invitees_names = employees.map(e => e.fullName).join(' و ');

        // A single variable containing the full, grammatically correct collective title.
        let collective_title = maleCollectiveTitle;
        if (femaleCollectiveTitle) {
            if (collective_title) {
                collective_title += femaleCollectiveTitle;
            } else {
                collective_title = femaleCollectiveTitle;
            }
        }
        
        return { city, maleCollectiveTitle, femaleCollectiveTitle, computedVar, combinedJobTitle, invitees_names, collective_title };
    }

    /**
     * Main function to generate and download Word documents.
     */
    generateWordBtn.addEventListener('click', () => {
        try {
            if (!wordTemplate) {
                alert('الرجاء تحميل قالب Word أولاً.');
                return;
            }
            if (invitedList.length === 0) {
                alert('لائحة المدعوين فارغة. الرجاء إضافة موظفين أولاً.');
                return;
            }
            if (typeof PizZip === 'undefined' || typeof docxtemplater === 'undefined') {
                alert('عذرًا، حدث خطأ أثناء تحميل مكتبات إنشاء الالمستندات.');
                return;
            }

            const groupedEmployees = invitedList.reduce((acc, emp) => {
                const key = `${emp.workLocation || 'N/A'}_${emp.division || 'N/A'}_${emp.city || 'N/A'}`;
                if (!acc[key]) {
                    acc[key] = {
                        workLocation: emp.workLocation,
                        division: emp.division,
                        city: emp.city,
                        employees: []
                    };
                }
                acc[key].employees.push(emp);
                return acc;
            }, {});

            for (const key in groupedEmployees) {
                const group = groupedEmployees[key];
                const { workLocation, division, employees } = group;

                const responsiblePerson = employeeData.find(e => {
                    if (e.workLocation !== workLocation || e.division !== division || e.city !== employees[0].city) return false;
                    return isResponsible(e);
                });

                let responsibilityLine = '';
                if (responsiblePerson) {
                    const genderPrefix = responsiblePerson.gender === 'السيدة' ? 'السيدة' : 'السيد';
                    responsibilityLine = `${genderPrefix} ${responsiblePerson.jobTitle}`;
                }
                
                const responsiblePersonId = responsiblePerson ? responsiblePerson.employeeId : null;
                const filteredEmployees = employees.filter(e => e.employeeId !== responsiblePersonId);
                
                const processedData = processGroupData(filteredEmployees);

                const zip = new PizZip(wordTemplate);
                const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

                doc.setData({
                    workLocation: workLocation,
                    division: division,
                    employees: filteredEmployees,
                    responsibilityLine: responsibilityLine,
                    ...processedData
                });

                doc.render();

                const out = doc.getZip().generate({
                    type: 'blob',
                    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                });
                
                const url = URL.createObjectURL(out);
                const a = document.createElement('a');
                a.href = url;
                const city = processedData.city || 'N/A';
                a.download = `${city}-${workLocation}-${division}.docx`;
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                a.remove();
            }
            alert('تم إنشاء الالمستندات بنجاح!');
        } catch (error) {
            console.error('An unexpected error occurred during Word generation:', error);
            alert(`حدث خطأ غير متوقع أثناء إنشاء الملف: ${error.message}`);
        }
    });

    // --- Initial Application Load ---
    loadEmployeeData();
    loadInvitedList();
});
