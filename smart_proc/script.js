// ============================================================
// PROCURESMART
// SMART PROCUREMENT MANAGEMENT SYSTEM
// ============================================================


// ============================================================
// GLOBAL VARIABLES
// ============================================================

let procurementCentres = [];

let selectedCentre = null;

let userLatitude = null;

let userLongitude = null;



// ============================================================
// PAGE LOAD
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log("ProcureSmart started.");

        const bookingForm =
            document.getElementById("bookingForm");

        if (!bookingForm) {
            return;
        }

        initializeFarmerPage();

    }
);



// ============================================================
// INITIALIZE FARMER PAGE
// ============================================================

function initializeFarmerPage() {

    const produce =
        document.getElementById("produce");

    const quantity =
        document.getElementById("quantity");

    const date =
        document.getElementById("date");

    const centre =
        document.getElementById("centre");

    const gpsButton =
        document.getElementById("gpsButton");

    const bookingForm =
        document.getElementById("bookingForm");


    // --------------------------------------------------------
    // Load Excel database
    // --------------------------------------------------------

    loadProcurementData();


    // --------------------------------------------------------
    // GPS
    // --------------------------------------------------------

    gpsButton.addEventListener(
        "click",
        detectGPS
    );


    // --------------------------------------------------------
    // Crop changes
    // --------------------------------------------------------

    produce.addEventListener(
        "change",
        updateRecommendations
    );


    // --------------------------------------------------------
    // Quantity changes
    // --------------------------------------------------------

    quantity.addEventListener(
        "input",
        updateRecommendations
    );


    // --------------------------------------------------------
    // Date changes
    // --------------------------------------------------------

    date.addEventListener(
        "change",
        function () {

            updateRecommendations();

            if (selectedCentre) {

                loadAvailableSlots(
                    selectedCentre
                );

            }

        }
    );


    // --------------------------------------------------------
    // Centre manually changed
    // --------------------------------------------------------

    centre.addEventListener(
        "change",
        function () {

            const centreId =
                centre.value;

            selectedCentre =
                procurementCentres.find(
                    function (item) {

                        return String(
                            item.centre_id
                        ) === String(
                            centreId
                        );

                    }
                );


            if (selectedCentre) {

                showCentreInformation(
                    selectedCentre
                );

                loadAvailableSlots(
                    selectedCentre
                );

            }

        }
    );


    // --------------------------------------------------------
    // Submit
    // --------------------------------------------------------

    bookingForm.addEventListener(
        "submit",
        bookSlot
    );

}



// ============================================================
// NORMALIZE TEXT
// ============================================================

function normalizeText(value) {

    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");

}



// ============================================================
// LOAD EXCEL DATABASE
// ============================================================

async function loadProcurementData() {

    const status =
        document.getElementById(
            "databaseStatus"
        );


    try {

        status.textContent =
            "🔄 Loading procurement centre database...";


        const response =
            await fetch(
                "proc_data.xlsx"
            );


        if (!response.ok) {

            throw new Error(
                "proc_data.xlsx could not be loaded."
            );

        }


        const excelData =
            await response.arrayBuffer();


        if (
            typeof XLSX ===
            "undefined"
        ) {

            throw new Error(
                "XLSX library was not loaded."
            );

        }


        const workbook =
            XLSX.read(
                excelData,
                {
                    type: "array",
                    cellDates: true
                }
            );


        const sheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[
                sheetName
            ];


        procurementCentres =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ""
                }
            );


        // ----------------------------------------------------
        // Remove completely empty rows
        // ----------------------------------------------------

        procurementCentres =
            procurementCentres.filter(
                function (centre) {

                    return (
                        centre.centre_id !== "" &&
                        centre.centre_id !== null &&
                        centre.centre_id !== undefined
                    );

                }
            );


        console.log(
            "Database loaded:",
            procurementCentres.length
        );


        console.table(
            procurementCentres.slice(
                0,
                10
            )
        );


        if (
            procurementCentres.length ===
            0
        ) {

            throw new Error(
                "No procurement centres found."
            );

        }


        status.textContent =
            "✅ Database loaded successfully: " +
            procurementCentres.length +
            " centres";


        setupDate();


        // ----------------------------------------------------
        // If GPS already detected, recommend
        // ----------------------------------------------------

        updateRecommendations();

    }

    catch (error) {

        console.error(
            "Database error:",
            error
        );


        status.textContent =
            "❌ Could not load procurement database.";


        alert(
            "Could not load proc_data.xlsx.\n\n" +
            "Make sure proc_data.xlsx is in the same folder " +
            "as farmer.html and run the project using Live Server."
        );

    }

}



// ============================================================
// SET DATE
// ============================================================

function setupDate() {

    const dateInput =
        document.getElementById(
            "date"
        );


    if (!dateInput) {
        return;
    }


    const today =
        new Date();


    const todayString =
        formatDateForInput(
            today
        );


    dateInput.min =
        todayString;


    // --------------------------------------------------------
    // Find first date available in Excel
    // --------------------------------------------------------

    let firstDatabaseDate = null;


    for (
        let i = 0;
        i < procurementCentres.length;
        i++
    ) {

        const parsed =
            parseExcelDate(
                procurementCentres[i].slot_date
            );


        if (parsed) {

            firstDatabaseDate =
                parsed;

            break;

        }

    }


    // --------------------------------------------------------
    // Use database date if valid
    // --------------------------------------------------------

    if (firstDatabaseDate) {

        const databaseDateString =
            formatDateForInput(
                firstDatabaseDate
            );


        if (
            databaseDateString >=
            todayString
        ) {

            dateInput.value =
                databaseDateString;

        }

        else {

            dateInput.value =
                todayString;

        }

    }

    else {

        dateInput.value =
            todayString;

    }

}



// ============================================================
// FORMAT DATE FOR HTML INPUT
// ============================================================

function formatDateForInput(date) {

    const year =
        date.getFullYear();


    const month =
        String(
            date.getMonth() + 1
        ).padStart(
            2,
            "0"
        );


    const day =
        String(
            date.getDate()
        ).padStart(
            2,
            "0"
        );


    return (
        year +
        "-" +
        month +
        "-" +
        day
    );

}



// ============================================================
// PARSE EXCEL DATE
// ============================================================

function parseExcelDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    // --------------------------------------------------------
    // JavaScript Date
    // --------------------------------------------------------

    if (
        value instanceof Date &&
        !isNaN(value.getTime())
    ) {

        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );

    }


    // --------------------------------------------------------
    // Excel serial number
    // --------------------------------------------------------

    if (
        typeof value === "number" &&
        isFinite(value)
    ) {

        const excelEpoch =
            new Date(
                1899,
                11,
                30
            );


        const date =
            new Date(
                excelEpoch.getTime() +
                value * 86400000
            );


        return new Date(
            date.getFullYear(),
            date.getMonth(),
            date.getDate()
        );

    }


    // --------------------------------------------------------
    // String
    // --------------------------------------------------------

    const text =
        String(value)
            .trim();


    // DD-MM-YYYY
    let match =
        text.match(
            /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/
        );


    if (match) {

        return new Date(
            Number(match[3]),
            Number(match[2]) - 1,
            Number(match[1])
        );

    }


    // YYYY-MM-DD
    match =
        text.match(
            /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/
        );


    if (match) {

        return new Date(
            Number(match[1]),
            Number(match[2]) - 1,
            Number(match[3])
        );

    }


    const parsed =
        new Date(text);


    if (
        !isNaN(
            parsed.getTime()
        )
    ) {

        return new Date(
            parsed.getFullYear(),
            parsed.getMonth(),
            parsed.getDate()
        );

    }


    return null;

}



// ============================================================
// GPS DETECTION
// ============================================================

function detectGPS() {

    const gpsStatus =
        document.getElementById(
            "gpsStatus"
        );


    const gpsButton =
        document.getElementById(
            "gpsButton"
        );


    if (
        !navigator.geolocation
    ) {

        gpsStatus.innerHTML =
            "❌ GPS is not supported by this browser.";

        return;

    }


    gpsButton.disabled =
        true;


    gpsButton.innerText =
        "📍 Detecting location...";


    gpsStatus.innerHTML =
        "🔄 Getting your GPS location...";


    navigator.geolocation.getCurrentPosition(

        function (position) {

            userLatitude =
                position.coords.latitude;


            userLongitude =
                position.coords.longitude;


            console.log(
                "User Latitude:",
                userLatitude
            );


            console.log(
                "User Longitude:",
                userLongitude
            );


            gpsStatus.innerHTML =

                "✅ GPS location detected successfully." +

                "<br><br>" +

                "Latitude: <strong>" +
                userLatitude.toFixed(6) +
                "</strong>" +

                "<br>" +

                "Longitude: <strong>" +
                userLongitude.toFixed(6) +
                "</strong>";


            gpsButton.disabled =
                false;


            gpsButton.innerText =
                "📍 Detect My Location Again";


            // ------------------------------------------------
            // Now find centre
            // ------------------------------------------------

            updateRecommendations();

        },


        function (error) {

            console.error(
                "GPS Error:",
                error
            );


            gpsButton.disabled =
                false;


            gpsButton.innerText =
                "📍 Detect My Location Again";


            if (
                error.code ===
                error.PERMISSION_DENIED
            ) {

                gpsStatus.innerHTML =
                    "❌ Location permission denied. Please allow location access.";

            }

            else if (
                error.code ===
                error.POSITION_UNAVAILABLE
            ) {

                gpsStatus.innerHTML =
                    "❌ GPS location is unavailable.";

            }

            else if (
                error.code ===
                error.TIMEOUT
            ) {

                gpsStatus.innerHTML =
                    "❌ GPS request timed out. Try again.";

            }

            else {

                gpsStatus.innerHTML =
                    "❌ Unable to detect your location.";

            }

        },


        {
            enableHighAccuracy: true,
            timeout: 15000,
            maximumAge: 0
        }

    );

}



// ============================================================
// CROP MATCHING
// ============================================================

function cropsMatch(
    farmerCrop,
    centreCrop
) {

    farmerCrop =
        normalizeText(
            farmerCrop
        );


    centreCrop =
        normalizeText(
            centreCrop
        );


    if (
        !farmerCrop ||
        !centreCrop
    ) {

        return false;

    }


    // Rice / Paddy

    if (
        farmerCrop === "rice" ||
        farmerCrop === "paddy"
    ) {

        return (
            centreCrop === "rice" ||
            centreCrop === "paddy"
        );

    }


    // Maize / Corn

    if (
        farmerCrop === "maize" ||
        farmerCrop === "corn"
    ) {

        return (
            centreCrop === "maize" ||
            centreCrop === "corn"
        );

    }


    // General exact match

    return (
        farmerCrop ===
        centreCrop
    );

}



// ============================================================
// HAVERSINE DISTANCE
// ============================================================

function calculateDistance(
    lat1,
    lon1,
    lat2,
    lon2
) {

    const R =
        6371;


    const dLat =
        degreesToRadians(
            lat2 - lat1
        );


    const dLon =
        degreesToRadians(
            lon2 - lon1
        );


    const a =
        Math.sin(dLat / 2) *
        Math.sin(dLat / 2) +

        Math.cos(
            degreesToRadians(lat1)
        ) *

        Math.cos(
            degreesToRadians(lat2)
        ) *

        Math.sin(dLon / 2) *
        Math.sin(dLon / 2);


    const c =
        2 *
        Math.atan2(
            Math.sqrt(a),
            Math.sqrt(1 - a)
        );


    return R * c;

}



// ============================================================
// DEGREES TO RADIANS
// ============================================================

function degreesToRadians(
    degrees
) {

    return (
        degrees *
        Math.PI /
        180
    );

}



// ============================================================
// GET SELECTED DATE
// ============================================================

function getSelectedDate() {

    const dateInput =
        document.getElementById(
            "date"
        );


    if (!dateInput) {
        return null;
    }


    if (!dateInput.value) {
        return null;
    }


    return dateInput.value;

}



// ============================================================
// CHECK CENTRE DATE
// ============================================================

function centreDateMatches(
    centre,
    selectedDate
) {

    if (!selectedDate) {

        return true;

    }


    const centreDate =
        parseExcelDate(
            centre.slot_date
        );


    if (!centreDate) {

        return false;

    }


    return (
        formatDateForInput(
            centreDate
        ) ===
        selectedDate
    );

}



// ============================================================
// GET SUITABLE CENTRES
// ============================================================

function getSuitableCentres() {

    const produce =
        document.getElementById(
            "produce"
        ).value;


    const quantity =
        Number(
            document.getElementById(
                "quantity"
            ).value
        );


    const selectedDate =
        getSelectedDate();


    if (
        !produce
    ) {

        return [];

    }


    const results =
        procurementCentres.filter(
            function (centre) {

                // --------------------------------------------
                // Crop
                // --------------------------------------------

                const cropOK =
                    cropsMatch(
                        produce,
                        centre.crop
                    );


                if (!cropOK) {
                    return false;
                }


                // --------------------------------------------
                // Date
                // --------------------------------------------

                const dateOK =
                    centreDateMatches(
                        centre,
                        selectedDate
                    );


                if (!dateOK) {
                    return false;
                }


                // --------------------------------------------
                // Capacity
                // --------------------------------------------

                const availableCapacity =
                    Number(
                        centre.available_capacity_quintals
                    ) || 0;


                if (
                    quantity > 0 &&
                    availableCapacity <
                    quantity
                ) {

                    return false;

                }


                // --------------------------------------------
                // Available slots
                // --------------------------------------------

                const availableSlots =
                    Number(
                        centre.available_slots
                    ) || 0;


                if (
                    availableSlots <= 0
                ) {

                    return false;

                }


                // --------------------------------------------
                // GPS coordinates
                // --------------------------------------------

                const centreLat =
                    Number(
                        centre.latitude
                    );


                const centreLon =
                    Number(
                        centre.longitude
                    );


                if (
                    !isFinite(centreLat) ||
                    !isFinite(centreLon)
                ) {

                    return false;

                }


                return true;

            }
        );


    return results;

}



// ============================================================
// SMART SCORE
// ============================================================

function calculateSmartScore(
    centre
) {

    let score = 0;


    // --------------------------------------------------------
    // DISTANCE SCORE
    // --------------------------------------------------------

    const distance =
        centre.distance;


    if (
        distance <= 5
    ) {

        score += 40;

    }

    else if (
        distance <= 10
    ) {

        score += 35;

    }

    else if (
        distance <= 20
    ) {

        score += 28;

    }

    else if (
        distance <= 40
    ) {

        score += 20;

    }

    else {

        score += 10;

    }


    // --------------------------------------------------------
    // CAPACITY SCORE
    // --------------------------------------------------------

    const availableCapacity =
        Number(
            centre.available_capacity_quintals
        ) || 0;


    const totalCapacity =
        Number(
            centre.capacity_quintals
        ) || 1;


    const capacityPercentage =
        (
            availableCapacity /
            totalCapacity
        ) * 100;


    if (
        capacityPercentage >= 70
    ) {

        score += 25;

    }

    else if (
        capacityPercentage >= 50
    ) {

        score += 22;

    }

    else if (
        capacityPercentage >= 30
    ) {

        score += 18;

    }

    else if (
        capacityPercentage >= 15
    ) {

        score += 13;

    }

    else {

        score += 8;

    }


    // --------------------------------------------------------
    // SLOT SCORE
    // --------------------------------------------------------

    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    const totalSlots =
        Number(
            centre.total_slots
        ) || 1;


    const slotPercentage =
        (
            availableSlots /
            totalSlots
        ) * 100;


    if (
        slotPercentage >= 70
    ) {

        score += 20;

    }

    else if (
        slotPercentage >= 50
    ) {

        score += 17;

    }

    else if (
        slotPercentage >= 30
    ) {

        score += 14;

    }

    else if (
        slotPercentage >= 15
    ) {

        score += 10;

    }

    else {

        score += 6;

    }


    // --------------------------------------------------------
    // AVAILABILITY STATUS
    // --------------------------------------------------------

    const status =
        normalizeText(
            centre.availability_status
        );


    if (
        status === "available"
    ) {

        score += 15;

    }


    return Math.min(
        Math.round(score),
        100
    );

}



// ============================================================
// UPDATE RECOMMENDATIONS
// ============================================================

function updateRecommendations() {

    const centreSelect =
        document.getElementById(
            "centre"
        );


    const recommendationBox =
        document.getElementById(
            "recommendationBox"
        );


    const recommendationText =
        document.getElementById(
            "recommendationText"
        );


    if (!centreSelect) {
        return;
    }


    const produce =
        document.getElementById(
            "produce"
        ).value;


    // --------------------------------------------------------
    // No crop
    // --------------------------------------------------------

    if (!produce) {

        centreSelect.innerHTML =
            `
            <option value="">
                Select produce first
            </option>
            `;


        recommendationBox.style.display =
            "none";


        selectedCentre =
            null;


        clearSlots();


        return;

    }


    // --------------------------------------------------------
    // GPS required
    // --------------------------------------------------------

    if (
        userLatitude === null ||
        userLongitude === null
    ) {

        centreSelect.innerHTML =
            `
            <option value="">
                📍 Detect GPS location first
            </option>
            `;


        recommendationBox.style.display =
            "block";


        recommendationText.innerHTML =

            "📍 Please click " +
            "<strong>Detect My Location</strong> " +
            "so ProcureSmart can find the nearest suitable centre.";


        selectedCentre =
            null;


        clearSlots();


        return;

    }


    // --------------------------------------------------------
    // Find suitable centres
    // --------------------------------------------------------

    let suitableCentres =
        getSuitableCentres();


    // --------------------------------------------------------
    // Calculate distance
    // --------------------------------------------------------

    suitableCentres =
        suitableCentres.map(
            function (centre) {

                const distance =
                    calculateDistance(

                        userLatitude,
                        userLongitude,

                        Number(
                            centre.latitude
                        ),
                        Number(
                            centre.longitude
                        )

                    );


                const smartScore =
                    calculateSmartScore(
                        {
                            ...centre,
                            distance:
                                distance
                        }
                    );


                return {

                    ...centre,

                    distance:
                        distance,

                    smartScore:
                        smartScore

                };

            }
        );


    // --------------------------------------------------------
    // Sort
    // --------------------------------------------------------

    suitableCentres.sort(
        function (a, b) {

            // Smart score first
            if (
                b.smartScore !==
                a.smartScore
            ) {

                return (
                    b.smartScore -
                    a.smartScore
                );

            }


            // If score same, nearest first
            return (
                a.distance -
                b.distance
            );

        }
    );


    // --------------------------------------------------------
    // No suitable centre
    // --------------------------------------------------------

    if (
        suitableCentres.length ===
        0
    ) {

        centreSelect.innerHTML =
            `
            <option value="">
                ❌ No suitable centre found
            </option>
            `;


        recommendationBox.style.display =
            "block";


        recommendationText.innerHTML =

            "❌ No suitable procurement centre found." +

            "<br><br>" +

            "Possible reasons:" +

            "<br>• Selected crop is not available" +

            "<br>• Quantity exceeds available capacity" +

            "<br>• No slots are available" +

            "<br>• Selected date is not available in the database";


        selectedCentre =
            null;


        clearSlots();


        return;

    }


    // --------------------------------------------------------
    // Fill dropdown
    // --------------------------------------------------------

    centreSelect.innerHTML =
        "";


    suitableCentres.forEach(
        function (centre) {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                centre.centre_id;


            option.textContent =

                centre.centre_name +

                " | " +

                centre.distance.toFixed(2) +

                " km | " +

                "Slots: " +

                centre.available_slots +

                " | Score: " +

                centre.smartScore +
                "/100";


            centreSelect.appendChild(
                option
            );

        }
    );


    // --------------------------------------------------------
    // Best centre
    // --------------------------------------------------------

    const bestCentre =
        suitableCentres[0];


    selectedCentre =
        bestCentre;


    centreSelect.value =
        bestCentre.centre_id;


    // --------------------------------------------------------
    // Show information
    // --------------------------------------------------------

    showCentreInformation(
        bestCentre
    );


    // --------------------------------------------------------
    // Load exact Excel slots
    // --------------------------------------------------------

    loadAvailableSlots(
        bestCentre
    );

}



// ============================================================
// SHOW CENTRE INFORMATION
// ============================================================

function showCentreInformation(
    centre
) {

    const recommendationBox =
        document.getElementById(
            "recommendationBox"
        );


    const recommendationText =
        document.getElementById(
            "recommendationText"
        );


    const availableCapacity =
        Number(
            centre.available_capacity_quintals
        ) || 0;


    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    const bookedSlots =
        Number(
            centre.booked_slots
        ) || 0;


    const price =
        Number(
            centre.indicative_price_inr_per_quintal
        ) || 0;


    recommendationBox.style.display =
        "block";


    recommendationText.innerHTML =

        "<strong>🥇 Recommended Procurement Centre</strong>" +

        "<br><br>" +

        "🏢 <strong>" +
        centre.centre_name +
        "</strong>" +

        "<br><br>" +

        "📍 Location: " +
        centre.location +

        "<br>" +

        "🏛 District: " +
        centre.district +

        "<br>" +

        "🌾 Crop: " +
        centre.crop +

        "<br>" +

        "🗺 Distance: <strong>" +
        centre.distance.toFixed(2) +
        " km</strong>" +

        "<br>" +

        "📦 Available Capacity: <strong>" +
        availableCapacity +
        " quintals</strong>" +

        "<br>" +

        "🎟 Available Slots: <strong>" +
        availableSlots +
        "</strong>" +

        "<br>" +

        "👥 Booked Slots: " +
        bookedSlots +

        "<br>" +

        "💰 Indicative Price: " +

        (
            price > 0
            ? "₹" +
              price +
              " / quintal"
            : "Not available"
        ) +

        "<br><br>" +

        "⭐ <strong>Smart Score: " +
        centre.smartScore +
        "/100</strong>";

}



// ============================================================
// LOAD AVAILABLE SLOTS
// IMPORTANT: USES available_slot_time FROM EXCEL
// ============================================================

function loadAvailableSlots(
    centre
) {

    const slotSelect =
        document.getElementById(
            "slot"
        );


    if (!slotSelect) {
        return;
    }


    slotSelect.innerHTML =
        "";


    // --------------------------------------------------------
    // Check selected date
    // --------------------------------------------------------

    const selectedDate =
        getSelectedDate();


    if (!selectedDate) {

        slotSelect.innerHTML =
            `
            <option value="">
                Select a date first
            </option>
            `;

        return;

    }


    // --------------------------------------------------------
    // Check centre date
    // --------------------------------------------------------

    if (
        !centreDateMatches(
            centre,
            selectedDate
        )
    ) {

        slotSelect.innerHTML =
            `
            <option value="">
                No slots for selected date
            </option>
            `;

        return;

    }


    // --------------------------------------------------------
    // Available slot count
    // --------------------------------------------------------

    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    if (
        availableSlots <= 0
    ) {

        slotSelect.innerHTML =
            `
            <option value="">
                No slots available
            </option>
            `;

        return;

    }


    // --------------------------------------------------------
    // GET available_slot_time FROM XLSX
    // --------------------------------------------------------

    let slotString =
        String(
            centre.available_slot_time ||
            ""
        ).trim();


    if (!slotString) {

        slotSelect.innerHTML =
            `
            <option value="">
                No slot timings available
            </option>
            `;

        return;

    }


    // --------------------------------------------------------
    // Split timings
    //
    // Example:
    //
    // 13:00-14:00,
    // 14:00-15:00,
    // 15:00-16:00
    // --------------------------------------------------------

    let slots =
        slotString
            .split(",")
            .map(
                function (slot) {

                    return slot.trim();

                }
            )
            .filter(
                function (slot) {

                    return slot.length > 0;

                }
            );


    // --------------------------------------------------------
    // Limit according to available_slots
    // --------------------------------------------------------

    slots =
        slots.slice(
            0,
            availableSlots
        );


    // --------------------------------------------------------
    // Add options
    // --------------------------------------------------------

    slots.forEach(
        function (slot, index) {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                slot;


            option.textContent =
                "🕐 " + slot;


            slotSelect.appendChild(
                option
            );

        }
    );


    // --------------------------------------------------------
    // If nothing created
    // --------------------------------------------------------

    if (
        slots.length ===
        0
    ) {

        slotSelect.innerHTML =
            `
            <option value="">
                No slot timings available
            </option>
            `;

        return;

    }


    console.log(
        "Available timings for",
        centre.centre_name,
        ":",
        slots
    );

}



// ============================================================
// CLEAR SLOTS
// ============================================================

function clearSlots() {

    const slotSelect =
        document.getElementById(
            "slot"
        );


    if (!slotSelect) {
        return;
    }


    slotSelect.innerHTML =
        `
        <option value="">
            Select suitable centre first
        </option>
        `;

}



// ============================================================
// BOOK SLOT
// ============================================================

function bookSlot(
    event
) {

    event.preventDefault();


    // --------------------------------------------------------
    // Get values
    // --------------------------------------------------------

    const farmerName =
        document.getElementById(
            "farmerName"
        ).value.trim();


    const mobile =
        document.getElementById(
            "mobile"
        ).value.trim();


    const produce =
        document.getElementById(
            "produce"
        ).value;


    const quantity =
        Number(
            document.getElementById(
                "quantity"
            ).value
        );


    const date =
        document.getElementById(
            "date"
        ).value;


    const slot =
        document.getElementById(
            "slot"
        ).value;


    const centreId =
        document.getElementById(
            "centre"
        ).value;


    // --------------------------------------------------------
    // Validation
    // --------------------------------------------------------

    if (
        !farmerName
    ) {

        alert(
            "❌ Please enter farmer name."
        );

        return;

    }


    if (
        !/^[0-9]{10}$/.test(
            mobile
        )
    ) {

        alert(
            "❌ Please enter a valid 10-digit mobile number."
        );

        return;

    }


    if (
        !produce
    ) {

        alert(
            "❌ Please select produce."
        );

        return;

    }


    if (
        !quantity ||
        quantity <= 0
    ) {

        alert(
            "❌ Please enter a valid quantity in quintals."
        );

        return;

    }


    if (
        userLatitude === null ||
        userLongitude === null
    ) {

        alert(
            "❌ Please detect your GPS location first."
        );

        return;

    }


    if (
        !date
    ) {

        alert(
            "❌ Please select a date."
        );

        return;

    }


    if (
        !slot
    ) {

        alert(
            "❌ Please select an available time slot."
        );

        return;

    }


    // --------------------------------------------------------
    // Find centre
    // --------------------------------------------------------

    const centre =
        procurementCentres.find(
            function (item) {

                return String(
                    item.centre_id
                ) === String(
                    centreId
                );

            }
        );


    if (!centre) {

        alert(
            "❌ Please select a valid procurement centre."
        );

        return;

    }


    // --------------------------------------------------------
    // Capacity check
    // --------------------------------------------------------

    const availableCapacity =
        Number(
            centre.available_capacity_quintals
        ) || 0;


    if (
        quantity >
        availableCapacity
    ) {

        alert(

            "❌ Insufficient capacity.\n\n" +

            "Available capacity: " +
            availableCapacity +
            " quintals\n" +

            "Your quantity: " +
            quantity +
            " quintals"

        );

        return;

    }


    // --------------------------------------------------------
    // Slot check
    // --------------------------------------------------------

    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    if (
        availableSlots <= 0
    ) {

        alert(
            "❌ No slots are available."
        );

        return;

    }


    // --------------------------------------------------------
    // Verify selected slot is actually in Excel
    // --------------------------------------------------------

    const excelSlots =
        String(
            centre.available_slot_time ||
            ""
        )
        .split(",")
        .map(
            function (s) {

                return s.trim();

            }
        )
        .filter(
            function (s) {

                return s.length > 0;

            }
        );


    if (
        !excelSlots.includes(
            slot
        )
    ) {

        alert(
            "❌ Selected slot is not available."
        );

        return;

    }


    // --------------------------------------------------------
    // Generate token
    // --------------------------------------------------------

    const token =
        Math.floor(
            Math.random() * 900
        ) + 100;


    // --------------------------------------------------------
    // Calculate distance
    // --------------------------------------------------------

    const distance =
        calculateDistance(

            userLatitude,
            userLongitude,

            Number(
                centre.latitude
            ),
            Number(
                centre.longitude
            )

        );


    // --------------------------------------------------------
    // Smart score
    // --------------------------------------------------------

    const smartScore =
        calculateSmartScore(
            {
                ...centre,
                distance:
                    distance
            }
        );


    // --------------------------------------------------------
    // Booking object
    // --------------------------------------------------------

    const booking = {

        farmerName:
            farmerName,

        mobile:
            mobile,

        produce:
            produce,

        quantityQuintals:
            quantity,

        latitude:
            userLatitude,

        longitude:
            userLongitude,

        centreId:
            centre.centre_id,

        centreName:
            centre.centre_name,

        district:
            centre.district,

        centreLocation:
            centre.location,

        centreLatitude:
            Number(
                centre.latitude
            ),

        centreLongitude:
            Number(
                centre.longitude
            ),

        distanceKm:
            Number(
                distance.toFixed(2)
            ),

        date:
            date,

        slot:
            slot,

        token:
            token,

        smartScore:
            smartScore,

        status:
            "Slot Booked",

        bookedAt:
            new Date().toISOString()

    };


    // --------------------------------------------------------
    // Save booking
    // --------------------------------------------------------

    localStorage.setItem(
        "booking",
        JSON.stringify(
            booking
        )
    );


    console.log(
        "Booking created:",
        booking
    );


    // --------------------------------------------------------
    // Confirmation
    // --------------------------------------------------------

    alert(

        "✅ Procurement slot booked successfully!\n\n" +

        "🎟 Token: #" +
        token +

        "\n\n" +

        "🏢 Centre: " +
        centre.centre_name +

        "\n\n" +

        "📏 Distance: " +
        distance.toFixed(2) +
        " km" +

        "\n\n" +

        "📅 Date: " +
        date +

        "\n\n" +

        "🕐 Time: " +
        slot +

        "\n\n" +

        "📦 Quantity: " +
        quantity +
        " quintals"

    );


    // --------------------------------------------------------
    // Open queue page
    // --------------------------------------------------------

    window.location.href =
        "queue.html";

}