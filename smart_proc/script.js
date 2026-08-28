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

let procurementMap = null;


// ============================================================
// PAGE LOAD
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log("ProcureSmart started.");

        const bookingForm =
            document.getElementById("bookingForm");

        if (bookingForm) {
            initializeFarmerPage();
        }

        if (document.getElementById("indiaMap")) {
            initializeCentreMap();
        }

    }
);


// ============================================================
// INITIALIZE FARMER PAGE
// ============================================================

async function initializeCentreMap() {

    if (typeof L === "undefined" || typeof XLSX === "undefined") {
        return;
    }

    procurementMap = L.map("indiaMap", { zoomControl: true }).setView([22.5, 79], 4.6);

    L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
        attribution: "&copy; OpenStreetMap &copy; CARTO",
        maxZoom: 18
    }).addTo(procurementMap);

    try {
        const response = await fetch("proc_data.xlsx");
        const workbook = XLSX.read(await response.arrayBuffer(), { type: "array", cellDates: true });
        procurementCentres = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" })
            .filter(function (centre) {
                return centre.centre_id && Number.isFinite(Number(centre.latitude)) && Number.isFinite(Number(centre.longitude));
            });

        populateMapFilters();
        renderCentreMap();
    } catch (error) {
        document.getElementById("predictionList").innerHTML = "<p class=\"loading-copy\">Centre data could not be loaded.</p>";
        console.error("Map database error:", error);
    }
}

function centreLoadScore(centre) {
    const capacity = Number(centre.capacity_quintals) || 0;
    const booked = Number(centre.booked_quintals) || 0;
    const totalSlots = Number(centre.total_slots) || 0;
    const bookedSlots = Number(centre.booked_slots) || 0;
    const demand = Number(centre.requirement_quintals) || 0;
    const capacityLoad = capacity ? booked / capacity : 0;
    const slotLoad = totalSlots ? bookedSlots / totalSlots : 0;
    const demandLoad = capacity ? demand / capacity : 0;
    return Math.round(Math.min(100, (capacityLoad * 45) + (slotLoad * 35) + (demandLoad * 20)));
}

function centreLoadState(score, centre) {
    const status = normalizeText(centre.availability_status);
    if (status.includes("closed") || status.includes("unavailable")) return "quiet";
    if (score >= 65) return "busy";
    if (score >= 30) return "active";
    return "quiet";
}

function populateMapFilters() {
    const cropSelect = document.getElementById("mapCrop");
    const stateSelect = document.getElementById("mapState");
    [...new Set(procurementCentres.map(c => c.crop).filter(Boolean))].sort().forEach(function (crop) {
        cropSelect.appendChild(new Option(crop, crop));
    });
    [...new Set(procurementCentres.map(c => c.state).filter(Boolean))].sort().forEach(function (state) {
        stateSelect.appendChild(new Option(state, state));
    });
    cropSelect.addEventListener("change", renderCentreMap);
    stateSelect.addEventListener("change", renderCentreMap);
}

function renderCentreMap() {
    const crop = document.getElementById("mapCrop").value;
    const state = document.getElementById("mapState").value;
    const filtered = procurementCentres.filter(function (centre) {
        return (crop === "all" || normalizeText(centre.crop) === normalizeText(crop)) &&
            (state === "all" || centre.state === state);
    }).map(function (centre) {
        const score = centreLoadScore(centre);
        return Object.assign({}, centre, { loadScore: score, loadState: centreLoadState(score, centre) });
    });

    procurementMap.eachLayer(function (layer) { if (layer instanceof L.CircleMarker) procurementMap.removeLayer(layer); });
    filtered.forEach(function (centre) {
        const color = centre.loadState === "busy" ? "#e76f51" : centre.loadState === "active" ? "#e9b949" : "#2a9d8f";
        L.circleMarker([Number(centre.latitude), Number(centre.longitude)], { radius: centre.loadState === "busy" ? 10 : 7, color: "#fff", weight: 2, fillColor: color, fillOpacity: .9 })
            .bindPopup(mapPopup(centre)).addTo(procurementMap);
    });

    document.getElementById("mapCentreCount").textContent = filtered.length;
    document.getElementById("mapBusyCount").textContent = filtered.filter(c => c.loadState === "busy").length;
    document.getElementById("mapActiveCount").textContent = filtered.filter(c => c.loadState === "active").length;
    document.getElementById("predictionList").innerHTML = filtered.sort((a, b) => b.loadScore - a.loadScore).slice(0, 5).map(predictionCard).join("") || "<p class=\"loading-copy\">No centres match these filters.</p>";
}

function mapPopup(centre) {
    return `<div class="map-popup"><strong>${centre.centre_name || "Procurement Centre"}</strong><span>${centre.district || ""}, ${centre.state || ""}</span><b class="popup-${centre.loadState}">${centre.loadState.toUpperCase()} · ${centre.loadScore}% load</b><p>${centre.crop || "All crops"} · ${centre.available_slots || 0} slots open</p><p>${centre.available_capacity_quintals || 0} quintals available · Rs ${centre.indicative_price_inr_per_quintal || "-"}/quintal</p><small>${centre.location || "Location unavailable"}</small></div>`;
}

function predictionCard(centre) {
    return `<button class="prediction-card" type="button" onclick="focusCentre(${Number(centre.latitude)}, ${Number(centre.longitude)})"><span class="prediction-dot ${centre.loadState}"></span><span><strong>${centre.centre_name || "Procurement Centre"}</strong><small>${centre.district || centre.state || "India"} · ${centre.crop || "Mixed crops"}</small></span><b>${centre.loadScore}%</b></button>`;
}

function focusCentre(latitude, longitude) {
    procurementMap.setView([latitude, longitude], 8, { animate: true });
}

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

    if (gpsButton) {

        gpsButton.addEventListener(
            "click",
            detectGPS
        );

    }


    // --------------------------------------------------------
    // Crop changes
    // --------------------------------------------------------

    if (produce) {

        produce.addEventListener(
            "change",
            updateRecommendations
        );

    }


    // --------------------------------------------------------
    // Quantity changes
    // --------------------------------------------------------

    if (quantity) {

        quantity.addEventListener(
            "input",
            updateRecommendations
        );

    }


    // --------------------------------------------------------
    // Date changes
    // --------------------------------------------------------

    if (date) {

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

    }


    // --------------------------------------------------------
    // Centre manually changed
    // --------------------------------------------------------

    if (centre) {

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

    }


    // --------------------------------------------------------
    // Submit
    // --------------------------------------------------------

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


        if (status) {

            status.textContent =
                "✅ Database loaded successfully: " +
                procurementCentres.length +
                " centres";

        }


        setupDate();


        // ----------------------------------------------------
        // Refresh recommendations after database is loaded
        // ----------------------------------------------------

        updateRecommendations();

    }
    catch (error) {

        console.error(
            "Database error:",
            error
        );


        if (status) {

            status.textContent =
                "❌ Could not load procurement database.";

        }


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


    // Keep the current valid selection if there is one.
    // Otherwise use today's date.

    if (!dateInput.value) {

        dateInput.value =
            todayString;

    }

}


// ============================================================
// FORMAT DATE FOR INPUT
// ============================================================

function formatDateForInput(date) {

    if (!(date instanceof Date)) {
        return "";
    }


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


    if (
        value instanceof Date &&
        !isNaN(value)
    ) {

        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );

    }


    // Excel serial number

    if (
        typeof value === "number"
    ) {

        const parsed =
            XLSX.SSF.parse_date_code(
                value
            );


        if (parsed) {

            return new Date(
                parsed.y,
                parsed.m - 1,
                parsed.d
            );

        }

    }


    const text =
        String(value).trim();


    // YYYY-MM-DD

    const isoMatch =
        text.match(
            /^(\d{4})-(\d{1,2})-(\d{1,2})$/
        );


    if (isoMatch) {

        return new Date(
            Number(isoMatch[1]),
            Number(isoMatch[2]) - 1,
            Number(isoMatch[3])
        );

    }


    // DD-MM-YYYY or DD/MM/YYYY

    const dmyMatch =
        text.match(
            /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/
        );


    if (dmyMatch) {

        return new Date(
            Number(dmyMatch[3]),
            Number(dmyMatch[2]) - 1,
            Number(dmyMatch[1])
        );

    }


    const parsedDate =
        new Date(text);


    if (!isNaN(parsedDate)) {

        return new Date(
            parsedDate.getFullYear(),
            parsedDate.getMonth(),
            parsedDate.getDate()
        );

    }


    return null;

}
// ============================================================
// GET SELECTED DATE
// ============================================================

function getSelectedDate() {

    const dateInput =
        document.getElementById(
            "date"
        );


    if (
        !dateInput ||
        !dateInput.value
    ) {

        return null;

    }


    const parts =
        dateInput.value.split(
            "-"
        );


    if (
        parts.length !== 3
    ) {

        return null;

    }


    return new Date(
        Number(parts[0]),
        Number(parts[1]) - 1,
        Number(parts[2])
    );

}


// ============================================================
// CENTRE DATE MATCH
// ============================================================

function centreDateMatches(
    centre,
    selectedDate
) {

    if (
        !selectedDate
    ) {

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
        centreDate.getFullYear() ===
            selectedDate.getFullYear() &&

        centreDate.getMonth() ===
            selectedDate.getMonth() &&

        centreDate.getDate() ===
            selectedDate.getDate()
    );

}


// ============================================================
// CROP MATCHING
// ============================================================

function cropsMatch(
    selectedCrop,
    centreCrop
) {

    const selected =
        normalizeText(
            selectedCrop
        );


    const available =
        normalizeText(
            centreCrop
        );


    if (
        !selected ||
        !available
    ) {

        return false;

    }


    // Exact match

    if (
        selected === available
    ) {

        return true;

    }


    // Handle multiple crops in one Excel cell

    const crops =
        available
            .split(
                /[,;/|]+/
            )
            .map(
                function (crop) {

                    return normalizeText(
                        crop
                    );

                }
            )
            .filter(
                function (crop) {

                    return crop !== "";

                }
            );


    if (
        crops.includes(
            selected
        )
    ) {

        return true;

    }


    // Common crop aliases

    const aliases = {

        "rice": [
            "paddy",
            "rice"
        ],

        "paddy": [
            "rice",
            "paddy"
        ],

        "maize": [
            "corn",
            "maize"
        ],

        "corn": [
            "maize",
            "corn"
        ],

        "pulses": [
            "pulse",
            "pulses",
            "dal"
        ]

    };


    if (aliases[selected]) {

        return aliases[selected]
            .some(
                function (alias) {

                    return (
                        available === alias ||
                        crops.includes(alias)
                    );

                }
            );

    }


    return false;

}


// ============================================================
// CALCULATE DISTANCE
// HAVERSINE FORMULA
// ============================================================

function calculateDistance(
    lat1,
    lon1,
    lat2,
    lon2
) {

    const radius =
        6371;


    const toRadians =
        function (value) {

            return (
                value *
                Math.PI
            ) / 180;

        };


    const deltaLat =
        toRadians(
            lat2 - lat1
        );


    const deltaLon =
        toRadians(
            lon2 - lon1
        );


    const a =
        Math.sin(
            deltaLat / 2
        ) *
        Math.sin(
            deltaLat / 2
        ) +

        Math.cos(
            toRadians(lat1)
        ) *

        Math.cos(
            toRadians(lat2)
        ) *

        Math.sin(
            deltaLon / 2
        ) *
        Math.sin(
            deltaLon / 2
        );


    const c =
        2 *
        Math.atan2(
            Math.sqrt(a),
            Math.sqrt(1 - a)
        );


    return (
        radius * c
    );

}


// ============================================================
// GET SUITABLE CENTRES
// ============================================================

function getSuitableCentres() {

    const produceElement =
        document.getElementById(
            "produce"
        );


    const quantityElement =
        document.getElementById(
            "quantity"
        );


    const produce =
        produceElement ?
        produceElement.value :
        "";


    const quantity =
        Number(
            quantityElement ?
            quantityElement.value :
            0
        ) || 0;


    if (
        !produce ||
        procurementCentres.length === 0
    ) {

        return [];

    }


    // --------------------------------------------------------
    // First find centres matching the crop and basic
    // availability requirements.
    //
    // IMPORTANT:
    // The date is NOT used here as a hard rejection.
    // This prevents the old Excel demo date from causing
    // "No suitable centre found".
    // --------------------------------------------------------

    let suitableCentres =
        procurementCentres.filter(
            function (centre) {

                // Crop must match

                if (
                    !cropsMatch(
                        produce,
                        centre.crop
                    )
                ) {

                    return false;

                }


                // Centre must have sufficient capacity

                const capacity =
                    Number(
                        centre.available_capacity_quintals
                    ) || 0;


                if (
                    quantity > 0 &&
                    capacity < quantity
                ) {

                    return false;

                }


                // Centre must have available booking slots

                const availableSlots =
                    Number(
                        centre.available_slots
                    ) || 0;


                if (
                    availableSlots <= 0
                ) {

                    return false;

                }


                // Centre must have valid coordinates

                const latitude =
                    Number(
                        centre.latitude
                    );


                const longitude =
                    Number(
                        centre.longitude
                    );


                if (
                    !Number.isFinite(
                        latitude
                    ) ||

                    !Number.isFinite(
                        longitude
                    )
                ) {

                    return false;

                }


                return true;

            }
        );


    // --------------------------------------------------------
    // Prefer centres whose date matches the selected date.
    //
    // If the database has no matching date, KEEP the suitable
    // centres instead of returning an empty list.
    // --------------------------------------------------------

    const selectedDate =
        getSelectedDate();


    if (
        selectedDate &&
        suitableCentres.length > 0
    ) {

        const exactDateCentres =
            suitableCentres.filter(
                function (centre) {

                    return centreDateMatches(
                        centre,
                        selectedDate
                    );

                }
            );


        if (
            exactDateCentres.length > 0
        ) {

            suitableCentres =
                exactDateCentres;

        }

    }


    // --------------------------------------------------------
    // Calculate distance when user GPS is available
    // --------------------------------------------------------

    suitableCentres =
        suitableCentres.map(
            function (centre) {

                const copy =
                    Object.assign(
                        {},
                        centre
                    );


                if (
                    userLatitude !== null &&
                    userLongitude !== null
                ) {

                    copy.distance =
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

                }
                else {

                    copy.distance =
                        Infinity;

                }


                return copy;

            }
        );


    // --------------------------------------------------------
    // Sort by distance and availability
    // --------------------------------------------------------

    suitableCentres.sort(
        function (a, b) {

            if (
                a.distance !==
                b.distance
            ) {

                return (
                    a.distance -
                    b.distance
                );

            }


            const aSlots =
                Number(
                    a.available_slots
                ) || 0;


            const bSlots =
                Number(
                    b.available_slots
                ) || 0;


            return (
                bSlots -
                aSlots
            );

        }
    );


    return suitableCentres;

}


// ============================================================
// UPDATE RECOMMENDATIONS
// ============================================================

function updateRecommendations() {

    const centreSelect =
        document.getElementById(
            "centre"
        );


    if (!centreSelect) {
        return;
    }


    // Do not show a false "No suitable centre found"
    // while the Excel database is still loading.

    if (
        procurementCentres.length === 0
    ) {

        return;

    }


    const produce =
        document.getElementById(
            "produce"
        );


    if (
        !produce ||
        !produce.value
    ) {

        centreSelect.innerHTML =
            `
            <option value="">
                Select produce to find suitable centres
            </option>
            `;

        selectedCentre = null;

        clearCentreInformation();

        clearAvailableSlots();

        return;

    }


    const suitableCentres =
        getSuitableCentres();


    displaySuitableCentres(
        suitableCentres
    );

}
// ============================================================
// DISPLAY SUITABLE CENTRES
// ============================================================

function displaySuitableCentres(
    centres
) {

    const centreSelect =
        document.getElementById(
            "centre"
        );


    if (!centreSelect) {
        return;
    }


    centreSelect.innerHTML =
        "";


    // --------------------------------------------------------
    // No centre found
    // --------------------------------------------------------

    if (
        !centres ||
        centres.length === 0
    ) {

        centreSelect.innerHTML =
            `
            <option value="">
                No suitable centre found
            </option>
            `;

        selectedCentre = null;

        clearCentreInformation();

        clearAvailableSlots();

        return;

    }


    // --------------------------------------------------------
    // Default option
    // --------------------------------------------------------

    const defaultOption =
        document.createElement(
            "option"
        );


    defaultOption.value =
        "";


    defaultOption.textContent =
        "Select a procurement centre";


    centreSelect.appendChild(
        defaultOption
    );


    // --------------------------------------------------------
    // Add recommended centres
    // --------------------------------------------------------

    centres.forEach(
        function (centre, index) {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                centre.centre_id;


            let centreName =
                centre.centre_name ||
                "Procurement Centre";


            let location =
                centre.location ||
                centre.district ||
                "";


            let distanceText =
                "";


            if (
                Number.isFinite(
                    centre.distance
                ) &&
                centre.distance !== Infinity
            ) {

                distanceText =
                    " — " +
                    centre.distance.toFixed(1) +
                    " km away";

            }


            // Mark the nearest recommendation

            let recommendedText =
                "";


            if (
                index === 0
            ) {

                recommendedText =
                    "⭐ Recommended: ";

            }


            option.textContent =
                recommendedText +
                centreName +
                (
                    location ?
                    " (" + location + ")" :
                    ""
                ) +
                distanceText;


            centreSelect.appendChild(
                option
            );

        }
    );


    // --------------------------------------------------------
    // Automatically select the best centre
    // --------------------------------------------------------

    selectedCentre =
        centres[0];


    centreSelect.value =
        String(
            selectedCentre.centre_id
        );


    showCentreInformation(
        selectedCentre
    );


    loadAvailableSlots(
        selectedCentre
    );


    console.log(
        "Recommended centre:",
        selectedCentre
    );

}


// ============================================================
// SHOW CENTRE INFORMATION
// ============================================================

function showCentreInformation(
    centre
) {

    if (!centre) {
        return;
    }


    const infoBox =
        document.getElementById(
            "centreInfo"
        );


    if (!infoBox) {
        return;
    }


    const name =
        centre.centre_name ||
        "Not available";


    const location =
        centre.location ||
        centre.district ||
        "Not available";


    const capacity =
        Number(
            centre.available_capacity_quintals
        ) || 0;


    const slots =
        Number(
            centre.available_slots
        ) || 0;


    let distance =
        "";


    if (
        Number.isFinite(
            centre.distance
        ) &&
        centre.distance !== Infinity
    ) {

        distance =
            centre.distance.toFixed(2) +
            " km";

    }
    else {

        distance =
            "Location distance unavailable";

    }


    infoBox.innerHTML =
        `
        <div class="centre-details">

            <h3>
                ⭐ Recommended Centre
            </h3>

            <p>
                <strong>🏢 Centre:</strong>
                ${name}
            </p>

            <p>
                <strong>📍 Location:</strong>
                ${location}
            </p>

            <p>
                <strong>📏 Distance:</strong>
                ${distance}
            </p>

            <p>
                <strong>📦 Available Capacity:</strong>
                ${capacity} quintals
            </p>

            <p>
                <strong>🕐 Available Slots:</strong>
                ${slots}
            </p>

        </div>
        `;

}


// ============================================================
// CLEAR CENTRE INFORMATION
// ============================================================

function clearCentreInformation() {

    const infoBox =
        document.getElementById(
            "centreInfo"
        );


    if (infoBox) {

        infoBox.innerHTML =
            "";

    }

}


// ============================================================
// CLEAR AVAILABLE SLOTS
// ============================================================

function clearAvailableSlots() {

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
            Select a centre first
        </option>
        `;

}


// ============================================================
// LOAD AVAILABLE TIME SLOTS
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


    if (!centre) {

        clearAvailableSlots();

        return;

    }


    // --------------------------------------------------------
    // Check centre availability
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
    // Read timings from database
    //
    // We deliberately do not reject timings because of the
    // selected date. The Excel dataset may contain a fixed
    // demonstration date, while the booking page uses today's
    // date.
    // --------------------------------------------------------

    const rawSlots =
        String(
            centre.available_slot_time ||
            centre.available_slots_time ||
            centre.slot_time ||
            ""
        ).trim();


    if (!rawSlots) {

        // Fallback demo timings if the database has a slot
        // count but no readable timing string.

        const fallbackSlots = [
            "09:00 AM – 10:00 AM",
            "10:00 AM – 11:00 AM",
            "11:00 AM – 12:00 PM",
            "12:00 PM – 01:00 PM",
            "02:00 PM – 03:00 PM",
            "03:00 PM – 04:00 PM"
        ];


        addSlotOptions(
            fallbackSlots.slice(
                0,
                Math.min(
                    availableSlots,
                    fallbackSlots.length
                )
            )
        );


        return;

    }


    let slots =
        rawSlots
            .split(
                /[,;|]+/
            )
            .map(
                function (slot) {

                    return slot.trim();

                }
            )
            .filter(
                function (slot) {

                    return (
                        slot.length > 0
                    );

                }
            );


    if (
        slots.length === 0
    ) {

        slotSelect.innerHTML =
            `
            <option value="">
                No slot timings available
            </option>
            `;

        return;

    }


    // Display only the number of slots currently available

    slots =
        slots.slice(
            0,
            availableSlots
        );


    addSlotOptions(
        slots
    );

}


// ============================================================
// ADD SLOT OPTIONS TO DROPDOWN
// ============================================================

function addSlotOptions(
    slots
) {

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
            Select available time slot
        </option>
        `;


    slots.forEach(
        function (slot) {

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

}
// ============================================================
// GPS LOCATION DETECTION
// ============================================================

function detectGPS() {

    const gpsStatus =
        document.getElementById(
            "gpsStatus"
        );


    if (
        !navigator.geolocation
    ) {

        if (gpsStatus) {

            gpsStatus.textContent =
                "❌ GPS location is not supported by your browser.";

        }

        return;

    }


    if (gpsStatus) {

        gpsStatus.textContent =
            "📍 Detecting your location...";

    }


    navigator.geolocation.getCurrentPosition(

        function (position) {

            userLatitude =
                position.coords.latitude;

            userLongitude =
                position.coords.longitude;


            console.log(
                "User GPS:",
                userLatitude,
                userLongitude
            );


            if (gpsStatus) {

                gpsStatus.textContent =
                    "✅ Location detected successfully.";

            }


            // Recalculate centre distances and recommendations

            updateRecommendations();

        },


        function (error) {

            console.error(
                "GPS error:",
                error
            );


            if (gpsStatus) {

                let message =
                    "❌ Unable to detect location.";

                if (
                    error.code ===
                    error.PERMISSION_DENIED
                ) {

                    message =
                        "❌ Location permission was denied.";

                }
                else if (
                    error.code ===
                    error.POSITION_UNAVAILABLE
                ) {

                    message =
                        "❌ Location information is unavailable.";

                }
                else if (
                    error.code ===
                    error.TIMEOUT
                ) {

                    message =
                        "❌ Location request timed out.";

                }


                gpsStatus.textContent =
                    message;

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
// GENERATE BOOKING TOKEN
// ============================================================

function generateToken() {

    const previousBooking =
        JSON.parse(
            localStorage.getItem(
                "booking"
            )
        );


    if (
        previousBooking &&
        previousBooking.token
    ) {

        const oldToken =
            Number(
                previousBooking.token
            );


        if (
            Number.isFinite(
                oldToken
            )
        ) {

            return (
                oldToken + 1
            );

        }

    }


    const randomNumber =
        Math.floor(
            100 +
            Math.random() *
            900
        );


    return randomNumber;

}


// ============================================================
// GENERATE BOOKING ID
// ============================================================

function generateBookingId() {

    const timestamp =
        Date.now()
            .toString()
            .slice(-6);


    const random =
        Math.floor(
            100 +
            Math.random() *
            900
        );


    return (
        "PS-" +
        timestamp +
        "-" +
        random
    );

}


// ============================================================
// BOOK PROCUREMENT SLOT
// ============================================================

function bookSlot(
    event
) {

    event.preventDefault();


    const farmerName =
        document.getElementById(
            "farmerName"
        );


    const produce =
        document.getElementById(
            "produce"
        );


    const quantity =
        document.getElementById(
            "quantity"
        );


    const date =
        document.getElementById(
            "date"
        );


    const centre =
        document.getElementById(
            "centre"
        );


    const slot =
        document.getElementById(
            "slot"
        );


    // --------------------------------------------------------
    // Basic validation
    // --------------------------------------------------------

    if (
        !farmerName ||
        !farmerName.value.trim()
    ) {

        alert(
            "Please enter the farmer name."
        );

        return;

    }


    if (
        !produce ||
        !produce.value
    ) {

        alert(
            "Please select the produce."
        );

        return;

    }


    const enteredQuantity =
        Number(
            quantity ?
            quantity.value :
            0
        );


    if (
        !enteredQuantity ||
        enteredQuantity <= 0
    ) {

        alert(
            "Please enter a valid quantity."
        );

        return;

    }


    if (
        !date ||
        !date.value
    ) {

        alert(
            "Please select a preferred date."
        );

        return;

    }


    if (
        !selectedCentre ||
        !centre ||
        !centre.value
    ) {

        alert(
            "Please select a procurement centre."
        );

        return;

    }


    if (
        !slot ||
        !slot.value
    ) {

        alert(
            "Please select an available time slot."
        );

        return;

    }


    // --------------------------------------------------------
    // Generate booking details
    // --------------------------------------------------------

    const token =
        generateToken();


    const bookingId =
        generateBookingId();


    const booking =
        {
            bookingId:
                bookingId,

            token:
                token,

            farmerName:
                farmerName.value.trim(),

            produce:
                produce.value,

            quantity:
                enteredQuantity,

            date:
                date.value,

            centreId:
                selectedCentre.centre_id,

            centreName:
                selectedCentre.centre_name ||
                "",

            centreLocation:
                selectedCentre.location ||
                selectedCentre.district ||
                "",

            slot:
                slot.value,

            status:
                "Confirmed",

            createdAt:
                new Date().toISOString(),

            latitude:
                userLatitude,

            longitude:
                userLongitude
        };


    // --------------------------------------------------------
    // Save booking for tracking page
    // --------------------------------------------------------

    localStorage.setItem(
        "booking",
        JSON.stringify(
            booking
        )
    );


    // --------------------------------------------------------
    // Save booking history
    // --------------------------------------------------------

    saveBookingToHistory(
        booking
    );


    console.log(
        "Booking created:",
        booking
    );


    // --------------------------------------------------------
    // Success message
    // --------------------------------------------------------

    alert(
        "🎉 Booking Confirmed!\n\n" +

        "Token: #" +
        token +
        "\n" +

        "Booking ID: " +
        bookingId +
        "\n\n" +

        "You can now track your procurement status."
    );


    // --------------------------------------------------------
    // Redirect if tracking page exists
    // --------------------------------------------------------

    window.location.href =
        "tracking.html";

}


// ============================================================
// SAVE BOOKING HISTORY
// ============================================================

function saveBookingToHistory(
    booking
) {

    let history =
        JSON.parse(
            localStorage.getItem(
                "bookingHistory"
            )
        );


    if (
        !Array.isArray(
            history
        )
    ) {

        history = [];

    }


    history.push(
        booking
    );


    localStorage.setItem(
        "bookingHistory",
        JSON.stringify(
            history
        )
    );

}


// ============================================================
// GET BOOKING HISTORY
// ============================================================

function getBookingHistory() {

    const history =
        JSON.parse(
            localStorage.getItem(
                "bookingHistory"
            )
        );


    if (
        Array.isArray(
            history
        )
    ) {

        return history;

    }


    return [];

}
// ============================================================
// UPDATE BOOKING STATUS
// ============================================================

function updateBookingStatus(
    newStatus
) {

    const booking =
        JSON.parse(
            localStorage.getItem(
                "booking"
            )
        );


    if (!booking) {

        return false;

    }


    booking.status =
        newStatus;


    booking.updatedAt =
        new Date().toISOString();


    localStorage.setItem(
        "booking",
        JSON.stringify(
            booking
        )
    );


    // Update the same booking in history

    let history =
        getBookingHistory();


    history =
        history.map(
            function (item) {

                if (
                    item.bookingId ===
                    booking.bookingId
                ) {

                    return booking;

                }

                return item;

            }
        );


    localStorage.setItem(
        "bookingHistory",
        JSON.stringify(
            history
        )
    );


    return true;

}


// ============================================================
// GET CURRENT BOOKING
// ============================================================

function getCurrentBooking() {

    try {

        const booking =
            JSON.parse(
                localStorage.getItem(
                    "booking"
                )
            );


        return booking || null;

    }
    catch (error) {

        console.error(
            "Could not read booking:",
            error
        );

        return null;

    }

}


// ============================================================
// GET BOOKING BY ID
// ============================================================

function getBookingById(
    bookingId
) {

    const history =
        getBookingHistory();


    return (
        history.find(
            function (booking) {

                return (
                    String(
                        booking.bookingId
                    ) ===
                    String(
                        bookingId
                    )
                );

            }
        ) || null
    );

}


// ============================================================
// FORMAT DATE FOR DISPLAY
// ============================================================

function formatDisplayDate(
    value
) {

    if (!value) {

        return "Not available";

    }


    const date =
        parseExcelDate(
            value
        );


    if (!date) {

        return String(value);

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


// ============================================================
// FORMAT BOOKING DATE
// ============================================================

function formatBookingDate(
    value
) {

    if (!value) {

        return "Not available";

    }


    const date =
        new Date(value);


    if (
        isNaN(
            date.getTime()
        )
    ) {

        return formatDisplayDate(
            value
        );

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


// ============================================================
// GET STATUS INFORMATION
// ============================================================

function getStatusInfo(
    status
) {

    const normalized =
        normalizeText(
            status
        );


    const statusMap =
        {

            "confirmed": {
                icon: "📅",
                title: "Booking Confirmed",
                description:
                    "Your procurement slot has been successfully reserved.",
                step: 1
            },

            "scheduled": {
                icon: "📅",
                title: "Booking Scheduled",
                description:
                    "Your procurement visit has been scheduled.",
                step: 1
            },

            "arrived": {
                icon: "📍",
                title: "Arrived at Centre",
                description:
                    "Your arrival has been registered at the procurement centre.",
                step: 2
            },

            "waiting": {
                icon: "🕐",
                title: "Waiting in Queue",
                description:
                    "Please wait for your token to be called.",
                step: 3
            },

            "processing": {
                icon: "⚙️",
                title: "Procurement in Progress",
                description:
                    "Your produce is currently being processed.",
                step: 4
            },

            "completed": {
                icon: "✅",
                title: "Procurement Completed",
                description:
                    "Your procurement process has been completed successfully.",
                step: 5
            },

            "cancelled": {
                icon: "❌",
                title: "Booking Cancelled",
                description:
                    "This procurement booking has been cancelled.",
                step: 0
            }

        };


    return (
        statusMap[normalized] ||
        statusMap["confirmed"]
    );

}


// ============================================================
// CALCULATE PEOPLE AHEAD
// ============================================================

function calculatePeopleAhead(
    booking
) {

    if (
        !booking ||
        !booking.token
    ) {

        return 0;

    }


    const currentToken =
        Number(
            localStorage.getItem(
                "currentServingToken"
            )
        ) || 35;


    const bookingToken =
        Number(
            booking.token
        );


    return Math.max(
        bookingToken -
        currentToken -
        1,
        0
    );

}


// ============================================================
// CALCULATE ESTIMATED WAIT
// ============================================================

function calculateEstimatedWait(
    booking
) {

    const peopleAhead =
        calculatePeopleAhead(
            booking
        );


    // Estimated average processing time:
    // 5 minutes per farmer

    return (
        peopleAhead * 5
    );

}


// ============================================================
// GET CURRENT SERVING TOKEN
// ============================================================

function getCurrentServingToken() {

    return (
        Number(
            localStorage.getItem(
                "currentServingToken"
            )
        ) || 35
    );

}


// ============================================================
// SET CURRENT SERVING TOKEN
// ============================================================

function setCurrentServingToken(
    token
) {

    const validToken =
        Number(token);


    if (
        !Number.isFinite(
            validToken
        )
    ) {

        return false;

    }


    localStorage.setItem(
        "currentServingToken",
        String(validToken)
    );


    return true;

}
// ============================================================
// GET BOOKING PROGRESS STEPS
// ============================================================

function getBookingProgress(
    booking
) {

    const statusInfo =
        getStatusInfo(
            booking ?
            booking.status :
            "confirmed"
        );


    const currentStep =
        statusInfo.step;


    return [

        {
            step: 1,
            title: "Booked",
            completed:
                currentStep >= 1
        },

        {
            step: 2,
            title: "Arrived",
            completed:
                currentStep >= 2
        },

        {
            step: 3,
            title: "In Queue",
            completed:
                currentStep >= 3
        },

        {
            step: 4,
            title: "Processing",
            completed:
                currentStep >= 4
        },

        {
            step: 5,
            title: "Completed",
            completed:
                currentStep >= 5
        }

    ];

}


// ============================================================
// CANCEL CURRENT BOOKING
// ============================================================

function cancelBooking() {

    const booking =
        getCurrentBooking();


    if (!booking) {

        alert(
            "No active booking found."
        );

        return;

    }


    const confirmed =
        confirm(
            "Are you sure you want to cancel this booking?"
        );


    if (!confirmed) {
        return;
    }


    updateBookingStatus(
        "Cancelled"
    );


    alert(
        "Your booking has been cancelled."
    );


    window.location.reload();

}


// ============================================================
// CLEAR CURRENT BOOKING
// ============================================================

function clearCurrentBooking() {

    localStorage.removeItem(
        "booking"
    );

}


// ============================================================
// GET RECOMMENDED CENTRE
// ============================================================

function getRecommendedCentre() {

    const centres =
        getSuitableCentres();


    if (
        !centres ||
        centres.length === 0
    ) {

        return null;

    }


    return centres[0];

}


// ============================================================
// DEBUG CENTRE FILTERING
// ============================================================

function debugCentreRecommendation() {

    const produceElement =
        document.getElementById(
            "produce"
        );


    const quantityElement =
        document.getElementById(
            "quantity"
        );


    console.group(
        "ProcureSmart Centre Recommendation Debug"
    );


    console.log(
        "Total database centres:",
        procurementCentres.length
    );


    console.log(
        "Selected produce:",
        produceElement ?
        produceElement.value :
        ""
    );


    console.log(
        "Selected quantity:",
        quantityElement ?
        quantityElement.value :
        ""
    );


    console.log(
        "Selected date:",
        getSelectedDate()
    );


    console.log(
        "User GPS:",
        userLatitude,
        userLongitude
    );


    const suitable =
        getSuitableCentres();


    console.log(
        "Suitable centres found:",
        suitable.length
    );


    console.table(
        suitable.slice(
            0,
            10
        )
    );


    console.groupEnd();


    return suitable;

}


// ============================================================
// VALIDATE DATABASE STRUCTURE
// ============================================================

function validateProcurementData() {

    if (
        !Array.isArray(
            procurementCentres
        )
    ) {

        return false;

    }


    if (
        procurementCentres.length === 0
    ) {

        return false;

    }


    return true;

}


// ============================================================
// REFRESH CENTRE RECOMMENDATION
// ============================================================

function refreshCentreRecommendation() {

    if (
        !validateProcurementData()
    ) {

        return;

    }


    updateRecommendations();

}


// ============================================================
// RESET BOOKING FORM
// ============================================================

function resetBookingForm() {

    const form =
        document.getElementById(
            "bookingForm"
        );


    if (form) {

        form.reset();

    }


    selectedCentre =
        null;


    clearCentreInformation();


    clearAvailableSlots();


    setupDate();

}


// ============================================================
// CHECK IF USER HAS ACTIVE BOOKING
// ============================================================

function hasActiveBooking() {

    const booking =
        getCurrentBooking();


    if (!booking) {
        return false;
    }


    const status =
        normalizeText(
            booking.status
        );


    return (
        status !== "completed" &&
        status !== "cancelled"
    );

}


// ============================================================
// GET ACTIVE BOOKING MESSAGE
// ============================================================

function getActiveBookingMessage() {

    const booking =
        getCurrentBooking();


    if (!booking) {

        return "";

    }


    if (
        !hasActiveBooking()
    ) {

        return "";

    }


    return (
        "You already have an active booking " +
        "(Token #" +
        booking.token +
        ")."
    );

}


// ============================================================
// PAGE-SAFE INITIALIZATION HELPERS
// ============================================================

// These functions allow script.js to be included on other pages
// without breaking those pages if booking-form elements are absent.

function isFarmerPage() {

    return Boolean(
        document.getElementById(
            "bookingForm"
        )
    );

}


function isTrackingPage() {

    return Boolean(
        document.getElementById(
            "trackingContainer"
        ) ||
        document.getElementById(
            "trackingStatus"
        ) ||
        document.getElementById(
            "token"
        )
    );

}


// ============================================================
// STORAGE SAFETY CHECK
// ============================================================

function storageAvailable() {

    try {

        const testKey =
            "__procureSmartTest__";


        localStorage.setItem(
            testKey,
            "1"
        );


        localStorage.removeItem(
            testKey
        );


        return true;

    }
    catch (error) {

        console.error(
            "Local storage unavailable:",
            error
        );

        return false;

    }

}


// ============================================================
// INITIAL APPLICATION CHECK
// ============================================================

if (
    storageAvailable()
) {

    console.log(
        "ProcureSmart local storage is ready."
    );

}
// ============================================================
// TRACKING PAGE SUPPORT
// ============================================================

// Return a clean summary of the current booking.
// This is useful for tracking.html or any future dashboard.

function getTrackingData() {

    const booking =
        getCurrentBooking();


    if (!booking) {

        return null;

    }


    const peopleAhead =
        calculatePeopleAhead(
            booking
        );


    const estimatedWait =
        calculateEstimatedWait(
            booking
        );


    const statusInfo =
        getStatusInfo(
            booking.status
        );


    return {

        booking:
            booking,

        currentServingToken:
            getCurrentServingToken(),

        peopleAhead:
            peopleAhead,

        estimatedWait:
            estimatedWait,

        statusInfo:
            statusInfo,

        progress:
            getBookingProgress(
                booking
            )

    };

}


// ============================================================
// UPDATE TRACKING PAGE ELEMENTS
// ============================================================

function updateTrackingPage() {

    const tracking =
        getTrackingData();


    if (!tracking) {

        return false;

    }


    const booking =
        tracking.booking;


    // Token

    const tokenElement =
        document.getElementById(
            "token"
        );


    if (tokenElement) {

        tokenElement.textContent =
            "#" + booking.token;

    }


    // Current serving token

    const currentTokenElement =
        document.getElementById(
            "currentToken"
        );


    if (currentTokenElement) {

        currentTokenElement.textContent =
            tracking.currentServingToken;

    }


    // People ahead

    const peopleAheadElement =
        document.getElementById(
            "peopleAhead"
        );


    if (peopleAheadElement) {

        peopleAheadElement.textContent =
            tracking.peopleAhead;

    }


    // Estimated wait

    const waitTimeElement =
        document.getElementById(
            "waitTime"
        );


    if (waitTimeElement) {

        waitTimeElement.textContent =
            tracking.estimatedWait;

    }


    // Queue status

    const queueStatusElement =
        document.getElementById(
            "queueStatus"
        );


    if (queueStatusElement) {

        if (
            tracking.peopleAhead === 0
        ) {

            queueStatusElement.textContent =
                "🟢 It is almost your turn. Please remain ready.";

        }
        else {

            queueStatusElement.textContent =
                "🟢 Queue is moving normally. " +
                tracking.peopleAhead +
                " farmer(s) are ahead of you.";

        }

    }


    return true;

}


// ============================================================
// AUTOMATIC TRACKING PAGE INITIALIZATION
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        if (
            isTrackingPage()
        ) {

            updateTrackingPage();

        }

    }
);


// ============================================================
// FINAL DEBUG MESSAGE
// ============================================================

console.log(
    "ProcureSmart script.js loaded successfully."
);
