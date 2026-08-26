// ============================================================
// PROCURESMART
// SMART PROCUREMENT MANAGEMENT SYSTEM
// GPS BASED SMART CENTRE RECOMMENDATION
// ============================================================


// ============================================================
// GLOBAL VARIABLES
// ============================================================

let procurementCentres = [];

let selectedRecommendation = null;

let farmerLatitude = null;

let farmerLongitude = null;


// ============================================================
// PAGE INITIALIZATION
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log("ProcureSmart started.");

        const bookingForm =
            document.getElementById("bookingForm");

        if (bookingForm) {

            startFarmerSystem();

        }

    }
);


// ============================================================
// START FARMER SYSTEM
// ============================================================

function startFarmerSystem() {

    console.log("Farmer portal loaded.");

    loadProcurementData();


    const produce =
        document.getElementById("produce");


    const quantity =
        document.getElementById("quantity");


    const centre =
        document.getElementById("centre");


    const bookingForm =
        document.getElementById("bookingForm");


    if (produce) {

        produce.addEventListener(
            "change",
            recommendCentre
        );

    }


    if (quantity) {

        quantity.addEventListener(
            "input",
            recommendCentre
        );

    }


    if (centre) {

        centre.addEventListener(
            "change",
            centreChanged
        );

    }


    if (bookingForm) {

        bookingForm.addEventListener(
            "submit",
            bookSlot
        );

    }

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

        if (status) {

            status.textContent =
                "🔄 Loading procurement centre database...";

        }


        const response =
            await fetch("proc_data.xlsx");


        if (!response.ok) {

            throw new Error(
                "proc_data.xlsx could not be loaded."
            );

        }


        const excelData =
            await response.arrayBuffer();


        if (
            typeof XLSX === "undefined"
        ) {

            throw new Error(
                "XLSX library is not loaded."
            );

        }


        const workbook =
            XLSX.read(
                excelData,
                {
                    type: "array"
                }
            );


        const sheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[sheetName];


        procurementCentres =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ""
                }
            );


        console.log(
            "Excel database loaded."
        );


        console.log(
            "Number of centres:",
            procurementCentres.length
        );


        console.table(
            procurementCentres
        );


        if (
            procurementCentres.length === 0
        ) {

            throw new Error(
                "Excel file contains no data."
            );

        }


        // ====================================================
        // CHECK LATITUDE AND LONGITUDE
        // ====================================================

        const firstCentre =
            procurementCentres[0];


        console.log(
            "First centre latitude:",
            firstCentre.latitude
        );


        console.log(
            "First centre longitude:",
            firstCentre.longitude
        );


        if (
            firstCentre.latitude === undefined ||
            firstCentre.longitude === undefined
        ) {

            console.warn(
                "Latitude or longitude columns were not found."
            );

        }


        if (status) {

            status.textContent =
                "✅ Database loaded successfully: " +
                procurementCentres.length +
                " procurement centres";

        }


        setMinimumDate();

    }


    catch (error) {

        console.error(
            "Database loading error:",
            error
        );


        if (status) {

            status.textContent =
                "❌ Could not load procurement database.";

        }


        alert(
            "Could not load proc_data.xlsx.\n\n" +
            "Make sure proc_data.xlsx is in the same folder " +
            "as farmer.html and use Live Server."
        );

    }

}


// ============================================================
// SET MINIMUM DATE
// ============================================================

function setMinimumDate() {

    const dateInput =
        document.getElementById("date");


    if (!dateInput) {

        return;

    }


    const today =
        new Date();


    const year =
        today.getFullYear();


    const month =
        String(
            today.getMonth() + 1
        ).padStart(2, "0");


    const day =
        String(
            today.getDate()
        ).padStart(2, "0");


    const todayString =
        year +
        "-" +
        month +
        "-" +
        day;


    dateInput.min =
        todayString;


    // Default date = tomorrow

    const tomorrow =
        new Date();


    tomorrow.setDate(
        tomorrow.getDate() + 1
    );


    const tomorrowYear =
        tomorrow.getFullYear();


    const tomorrowMonth =
        String(
            tomorrow.getMonth() + 1
        ).padStart(2, "0");


    const tomorrowDay =
        String(
            tomorrow.getDate()
        ).padStart(2, "0");


    dateInput.value =
        tomorrowYear +
        "-" +
        tomorrowMonth +
        "-" +
        tomorrowDay;

}


// ============================================================
// GPS LOCATION
// ============================================================

function getFarmerLocation() {

    const locationStatus =
        document.getElementById(
            "locationStatus"
        );


    const locationButton =
        document.getElementById(
            "locationButton"
        );


    if (
        !navigator.geolocation
    ) {

        locationStatus.innerHTML =
            "❌ Your browser does not support GPS location.";

        return;

    }


    locationStatus.innerHTML =
        "📍 Detecting your current location...";


    if (locationButton) {

        locationButton.disabled =
            true;

        locationButton.innerText =
            "📍 Detecting...";

    }


    navigator.geolocation.getCurrentPosition(

        function (position) {

            farmerLatitude =
                position.coords.latitude;


            farmerLongitude =
                position.coords.longitude;


            // Save in hidden fields

            document.getElementById(
                "latitude"
            ).value =
                farmerLatitude;


            document.getElementById(
                "longitude"
            ).value =
                farmerLongitude;


            console.log(
                "Farmer latitude:",
                farmerLatitude
            );


            console.log(
                "Farmer longitude:",
                farmerLongitude
            );


            // Display coordinates

            locationStatus.innerHTML =

                "✅ Location detected." +

                "<br>" +

                "Latitude: " +
                farmerLatitude.toFixed(6) +

                "<br>" +

                "Longitude: " +
                farmerLongitude.toFixed(6);


            if (locationButton) {

                locationButton.disabled =
                    false;

                locationButton.innerText =
                    "📍 Update My Location";

            }


            // Automatically find recommendation

            recommendCentre();

        },


        function (error) {

            console.error(
                "GPS error:",
                error
            );


            let message =
                "❌ Could not detect your location.";


            if (
                error.code ===
                error.PERMISSION_DENIED
            ) {

                message =
                    "❌ Location permission was denied. " +
                    "Please allow location access in your browser.";

            }


            else if (
                error.code ===
                error.POSITION_UNAVAILABLE
            ) {

                message =
                    "❌ Your location is currently unavailable.";

            }


            else if (
                error.code ===
                error.TIMEOUT
            ) {

                message =
                    "❌ Location request timed out. " +
                    "Please try again.";

            }


            locationStatus.innerHTML =
                message;


            if (locationButton) {

                locationButton.disabled =
                    false;

                locationButton.innerText =
                    "📍 Try Again";

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
        farmerCrop.includes("paddy") ||
        farmerCrop.includes("rice")
    ) {

        return (
            centreCrop.includes("paddy") ||
            centreCrop.includes("rice")
        );

    }


    // Wheat

    if (
        farmerCrop.includes("wheat")
    ) {

        return centreCrop.includes(
            "wheat"
        );

    }


    // Maize / Corn

    if (
        farmerCrop.includes("maize") ||
        farmerCrop.includes("corn")
    ) {

        return (
            centreCrop.includes("maize") ||
            centreCrop.includes("corn")
        );

    }


    // Mustard

    if (
        farmerCrop.includes("mustard")
    ) {

        return centreCrop.includes(
            "mustard"
        );

    }


    // Coconut

    if (
        farmerCrop.includes("coconut")
    ) {

        return centreCrop.includes(
            "coconut"
        );

    }


    // Groundnut

    if (
        farmerCrop.includes("groundnut")
    ) {

        return centreCrop.includes(
            "groundnut"
        );

    }


    return (
        centreCrop.includes(
            farmerCrop
        ) ||
        farmerCrop.includes(
            centreCrop
        )
    );

}


// ============================================================
// CONVERT DATABASE COORDINATE TO NUMBER
// ============================================================

function getCoordinate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    const number =
        Number(
            String(value)
                .trim()
        );


    if (
        !Number.isFinite(number)
    ) {

        return null;

    }


    return number;

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

    const earthRadiusKm =
        6371;


    const latitudeDifference =
        toRadians(
            lat2 - lat1
        );


    const longitudeDifference =
        toRadians(
            lon2 - lon1
        );


    const a =

        Math.sin(
            latitudeDifference / 2
        ) *
        Math.sin(
            latitudeDifference / 2
        )

        +

        Math.cos(
            toRadians(lat1)
        ) *

        Math.cos(
            toRadians(lat2)
        ) *

        Math.sin(
            longitudeDifference / 2
        ) *

        Math.sin(
            longitudeDifference / 2
        );


    const c =
        2 *
        Math.atan2(
            Math.sqrt(a),
            Math.sqrt(1 - a)
        );


    return (
        earthRadiusKm * c
    );

}


// ============================================================
// DEGREES TO RADIANS
// ============================================================

function toRadians(
    degrees
) {

    return (
        degrees *
        Math.PI /
        180
    );

}


// ============================================================
// FIND SUITABLE CENTRES
// ============================================================

function findSuitableCentres() {

    const produceInput =
        document.getElementById(
            "produce"
        );


    const quantityInput =
        document.getElementById(
            "quantity"
        );


    if (
        !produceInput ||
        !quantityInput
    ) {

        return [];

    }


    const farmerCrop =
        produceInput.value;


    const quantityQuintals =
        Number(
            quantityInput.value
        );


    if (
        !farmerCrop ||
        quantityQuintals <= 0
    ) {

        return [];

    }


    const suitableCentres =
        procurementCentres.filter(
            function (centre) {


                // --------------------------------------------
                // CROP
                // --------------------------------------------

                const cropMatches =
                    cropsMatch(
                        farmerCrop,
                        centre.crop
                    );


                // --------------------------------------------
                // CAPACITY
                // --------------------------------------------

                const availableCapacity =
                    Number(
                        centre.available_capacity_quintals
                    ) || 0;


                const capacityAvailable =
                    availableCapacity >=
                    quantityQuintals;


                // --------------------------------------------
                // SLOTS
                // --------------------------------------------

                const availableSlots =
                    Number(
                        centre.available_slots
                    ) || 0;


                const slotsAvailable =
                    availableSlots > 0;


                return (
                    cropMatches &&
                    capacityAvailable &&
                    slotsAvailable
                );

            }
        );


    return suitableCentres;

}


// ============================================================
// SMART SCORE
// ============================================================

function calculateSmartScore(
    centre,
    distanceKm
) {

    let score = 0;


    // ========================================================
    // DISTANCE SCORE
    // Maximum = 45
    // ========================================================

    if (
        distanceKm === null
    ) {

        score += 0;

    }

    else if (
        distanceKm <= 5
    ) {

        score += 45;

    }

    else if (
        distanceKm <= 10
    ) {

        score += 40;

    }

    else if (
        distanceKm <= 20
    ) {

        score += 35;

    }

    else if (
        distanceKm <= 30
    ) {

        score += 28;

    }

    else if (
        distanceKm <= 50
    ) {

        score += 20;

    }

    else if (
        distanceKm <= 100
    ) {

        score += 10;

    }

    else {

        score += 3;

    }


    // ========================================================
    // CAPACITY SCORE
    // Maximum = 25
    // ========================================================

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
        ) *
        100;


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

        score += 12;

    }

    else {

        score += 5;

    }


    // ========================================================
    // SLOT SCORE
    // Maximum = 20
    // ========================================================

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
        ) *
        100;


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

        score += 5;

    }


    // ========================================================
    // QUEUE SCORE
    // Maximum = 10
    // ========================================================

    const bookedSlots =
        Number(
            centre.booked_slots
        ) || 0;


    const queueLoad =
        (
            bookedSlots /
            totalSlots
        ) *
        100;


    if (
        queueLoad <= 20
    ) {

        score += 10;

    }

    else if (
        queueLoad <= 40
    ) {

        score += 8;

    }

    else if (
        queueLoad <= 60
    ) {

        score += 6;

    }

    else if (
        queueLoad <= 80
    ) {

        score += 4;

    }

    else {

        score += 2;

    }


    return Math.min(
        Math.round(score),
        100
    );

}


// ============================================================
// RECOMMEND CENTRE
// ============================================================

function recommendCentre() {

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


    // ========================================================
    // CHECK GPS
    // ========================================================

    if (
        farmerLatitude === null ||
        farmerLongitude === null
    ) {

        centreSelect.innerHTML =

            `<option value="">
                Detect your location first
            </option>`;


        if (recommendationBox) {

            recommendationBox.style.display =
                "block";

        }


        if (recommendationText) {

            recommendationText.innerHTML =

                `📍 Please click
                <strong>
                "Detect My Location"
                </strong>
                so the system can find the nearest procurement centre.`;

        }


        selectedRecommendation =
            null;


        return;

    }


    // ========================================================
    // FIND SUITABLE CENTRES
    // ========================================================

    const suitableCentres =
        findSuitableCentres();


    centreSelect.innerHTML =
        "";


    if (
        suitableCentres.length === 0
    ) {

        centreSelect.innerHTML =

            `<option value="">
                No suitable centre found
            </option>`;


        if (recommendationBox) {

            recommendationBox.style.display =
                "block";

        }


        if (recommendationText) {

            recommendationText.innerHTML =

                `❌ No suitable procurement centre found.

                <br><br>

                Please check:

                <br>
                • Produce selected

                <br>
                • Quantity in quintals

                <br>
                • Available capacity

                <br>
                • Available slots`;

        }


        selectedRecommendation =
            null;


        return;

    }


    // ========================================================
    // CALCULATE DISTANCE
    // ========================================================

    const scoredCentres =
        suitableCentres.map(
            function (centre) {


                const centreLat =
                    getCoordinate(
                        centre.latitude
                    );


                const centreLon =
                    getCoordinate(
                        centre.longitude
                    );


                let distanceKm =
                    null;


                if (
                    centreLat !== null &&
                    centreLon !== null
                ) {

                    distanceKm =
                        calculateDistance(
                            farmerLatitude,
                            farmerLongitude,
                            centreLat,
                            centreLon
                        );

                }


                const smartScore =
                    calculateSmartScore(
                        centre,
                        distanceKm
                    );


                return {

                    ...centre,

                    distanceKm:
                        distanceKm,

                    smartScore:
                        smartScore

                };

            }
        );


    // ========================================================
    // SORT
    // ========================================================

    scoredCentres.sort(
        function (a, b) {

            // First priority = smart score

            if (
                b.smartScore !==
                a.smartScore
            ) {

                return (
                    b.smartScore -
                    a.smartScore
                );

            }


            // Second priority = distance

            if (
                a.distanceKm !== null &&
                b.distanceKm !== null
            ) {

                return (
                    a.distanceKm -
                    b.distanceKm
                );

            }


            return 0;

        }
    );


    // ========================================================
    // ADD CENTRES TO DROPDOWN
    // ========================================================

    scoredCentres.forEach(
        function (centre) {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                centre.centre_id;


            let distanceText =
                "Distance unavailable";


            if (
                centre.distanceKm !== null
            ) {

                distanceText =
                    centre.distanceKm.toFixed(1) +
                    " km away";

            }


            option.textContent =

                centre.centre_name +

                " - " +

                distanceText +

                " - Score " +

                centre.smartScore +

                "/100";


            centreSelect.appendChild(
                option
            );

        }
    );


    // ========================================================
    // BEST CENTRE
    // ========================================================

    const bestCentre =
        scoredCentres[0];


    selectedRecommendation =
        bestCentre;


    centreSelect.value =
        bestCentre.centre_id;


    // ========================================================
    // INFORMATION
    // ========================================================

    const availableCapacity =
        Number(
            bestCentre.available_capacity_quintals
        ) || 0;


    const availableSlots =
        Number(
            bestCentre.available_slots
        ) || 0;


    const bookedSlots =
        Number(
            bestCentre.booked_slots
        ) || 0;


    const price =
        Number(
            bestCentre
                .indicative_price_inr_per_quintal
        ) || 0;


    let distanceText =
        "Distance unavailable";


    if (
        bestCentre.distanceKm !== null
    ) {

        distanceText =
            bestCentre.distanceKm.toFixed(2) +
            " km";

    }


    if (recommendationBox) {

        recommendationBox.style.display =
            "block";

    }


    if (recommendationText) {

        recommendationText.innerHTML =

            `

            <strong>
                🥇 Recommended Procurement Centre
            </strong>

            <br><br>

            🏢
            <strong>
                ${bestCentre.centre_name}
            </strong>

            <br>

            📍 Centre Location:
            ${bestCentre.location}

            <br>

            🧭 Distance from you:
            <strong>
                ${distanceText}
            </strong>

            <br>

            🏛 District:
            ${bestCentre.district}

            <br>

            🌾 Crop:
            ${bestCentre.crop}

            <br>

            📦 Available Capacity:
            ${availableCapacity}
            quintals

            <br>

            🎟 Available Slots:
            ${availableSlots}

            <br>

            👥 Booked Slots:
            ${bookedSlots}

            <br>

            💰 Indicative Price:
            ${
                price > 0
                ? "₹" + price +
                  " / quintal"
                : "Not available"
            }

            <br><br>

            ⭐
            <strong>
                Smart Score:
                ${bestCentre.smartScore}/100
            </strong>

            `;

    }


    // ========================================================
    // LOAD SLOTS
    // ========================================================

    loadAvailableSlots(
        bestCentre
    );

}


// ============================================================
// CENTRE CHANGED
// ============================================================

function centreChanged() {

    const centreId =
        document.getElementById(
            "centre"
        ).value;


    const centre =
        procurementCentres.find(
            function (item) {

                return String(
                    item.centre_id
                ) ===
                String(
                    centreId
                );

            }
        );


    if (centre) {

        selectedRecommendation =
            centre;


        loadAvailableSlots(
            centre
        );

    }

}


// ============================================================
// LOAD TIME SLOTS
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


    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    if (
        availableSlots <= 0
    ) {

        slotSelect.innerHTML =

            `<option value="">
                No slots available
            </option>`;

        return;

    }


    const operatingStart =
        String(
            centre.operating_start ||
            "08:00"
        );


    const operatingEnd =
        String(
            centre.operating_end ||
            "17:00"
        );


    const slotDuration =
        Number(
            centre.slot_duration_minutes
        ) || 60;


    let currentMinutes =
        convertTimeToMinutes(
            operatingStart
        );


    const endMinutes =
        convertTimeToMinutes(
            operatingEnd
        );


    let slotsCreated =
        0;


    while (
        currentMinutes <
        endMinutes &&
        slotsCreated <
        availableSlots
    ) {

        const time =
            convertMinutesToTime(
                currentMinutes
            );


        const option =
            document.createElement(
                "option"
            );


        option.value =
            time;


        option.textContent =
            time;


        slotSelect.appendChild(
            option
        );


        currentMinutes +=
            slotDuration;


        slotsCreated++;

    }

}


// ============================================================
// TIME TO MINUTES
// ============================================================

function convertTimeToMinutes(
    time
) {

    if (
        !time ||
        !time.includes(":")
    ) {

        return 0;

    }


    const parts =
        time.split(":");


    const hours =
        Number(
            parts[0]
        ) || 0;


    const minutes =
        Number(
            parts[1]
        ) || 0;


    return (
        hours * 60 +
        minutes
    );

}


// ============================================================
// MINUTES TO TIME
// ============================================================

function convertMinutesToTime(
    totalMinutes
) {

    const hours24 =
        Math.floor(
            totalMinutes / 60
        );


    const minutes =
        totalMinutes % 60;


    const period =
        hours24 >= 12
        ? "PM"
        : "AM";


    let hours12 =
        hours24 % 12;


    if (
        hours12 === 0
    ) {

        hours12 = 12;

    }


    return (

        String(
            hours12
        ).padStart(
            2,
            "0"
        )

        +

        ":" +

        String(
            minutes
        ).padStart(
            2,
            "0"
        )

        +

        " " +

        period

    );

}


// ============================================================
// BOOK SLOT
// ============================================================

function bookSlot(event) {

    event.preventDefault();


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


    const quantityQuintals =
        Number(
            document.getElementById(
                "quantity"
            ).value
        );


    const location =
        document.getElementById(
            "location"
        ).value.trim();


    const centreId =
        document.getElementById(
            "centre"
        ).value;


    const date =
        document.getElementById(
            "date"
        ).value;


    const slot =
        document.getElementById(
            "slot"
        ).value;


    // ========================================================
    // GPS VALIDATION
    // ========================================================

    if (
        farmerLatitude === null ||
        farmerLongitude === null
    ) {

        alert(
            "📍 Please detect your current location before booking."
        );

        return;

    }


    // ========================================================
    // FIND CENTRE
    // ========================================================

    const centre =
        procurementCentres.find(
            function (item) {

                return String(
                    item.centre_id
                ) ===
                String(
                    centreId
                );

            }
        );


    if (!centre) {

        alert(
            "❌ Please select a procurement centre."
        );

        return;

    }


    // ========================================================
    // CAPACITY
    // ========================================================

    const availableCapacity =
        Number(
            centre.available_capacity_quintals
        ) || 0;


    if (
        quantityQuintals >
        availableCapacity
    ) {

        alert(
            "❌ The selected centre does not have enough capacity."
        );

        return;

    }


    // ========================================================
    // SLOTS
    // ========================================================

    const availableSlots =
        Number(
            centre.available_slots
        ) || 0;


    if (
        availableSlots <= 0
    ) {

        alert(
            "❌ No slots are available at this centre."
        );

        return;

    }


    if (!slot) {

        alert(
            "❌ Please select a time slot."
        );

        return;

    }


    // ========================================================
    // DISTANCE
    // ========================================================

    const centreLat =
        getCoordinate(
            centre.latitude
        );


    const centreLon =
        getCoordinate(
            centre.longitude
        );


    let distanceKm =
        null;


    if (
        centreLat !== null &&
        centreLon !== null
    ) {

        distanceKm =
            calculateDistance(
                farmerLatitude,
                farmerLongitude,
                centreLat,
                centreLon
            );

    }


    // ========================================================
    // SMART SCORE
    // ========================================================

    const smartScore =
        calculateSmartScore(
            centre,
            distanceKm
        );


    // ========================================================
    // TOKEN
    // ========================================================

    const token =
        Math.floor(
            Math.random() * 900
        ) + 100;


    // ========================================================
    // BOOKING OBJECT
    // ========================================================

    const booking = {

        farmerName:
            farmerName,

        mobile:
            mobile,

        location:
            location,

        farmerLatitude:
            farmerLatitude,

        farmerLongitude:
            farmerLongitude,

        produce:
            produce,

        quantityQuintals:
            quantityQuintals,

        centreId:
            centre.centre_id,

        centreName:
            centre.centre_name,

        district:
            centre.district,

        centreLocation:
            centre.location,

        centreLatitude:
            centreLat,

        centreLongitude:
            centreLon,

        distanceKm:
            distanceKm,

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


    // ========================================================
    // SAVE BOOKING
    // ========================================================

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


    // ========================================================
    // CONFIRMATION
    // ========================================================

    let distanceMessage =
        "Distance unavailable";


    if (
        distanceKm !== null
    ) {

        distanceMessage =
            distanceKm.toFixed(2) +
            " km";

    }


    alert(

        "✅ Procurement slot booked successfully!" +

        "\n\n" +

        "Token: #" +
        token +

        "\n\nCentre: " +
        centre.centre_name +

        "\n\nDistance: " +
        distanceMessage +

        "\n\nDate: " +
        date +

        "\n\nTime: " +
        slot +

        "\n\nQuantity: " +
        quantityQuintals +
        " quintals"

    );


    // ========================================================
    // OPEN QUEUE
    // ========================================================

    window.location.href =
        "queue.html";

}