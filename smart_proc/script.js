document
    .getElementById("bookingForm")
    ?.addEventListener("submit", function(event) {

        event.preventDefault();

        const name =
            document.getElementById("farmerName").value;

        const produce =
            document.getElementById("produce").value;

        const quantity =
            document.getElementById("quantity").value;

        const centre =
            document.getElementById("centre").value;

        const date =
            document.getElementById("date").value;

        const slot =
            document.getElementById("slot").value;


        // Generate demo token

        const token =
            Math.floor(Math.random() * 50) + 1;


        // Save booking in browser

        const booking = {

            name: name,

            produce: produce,

            quantity: quantity,

            centre: centre,

            date: date,

            slot: slot,

            token: token,

            status: "Slot Booked"

        };


        localStorage.setItem(
            "booking",
            JSON.stringify(booking)
        );


        // Show confirmation

        alert(
            "✅ Slot Booked Successfully!\n\n" +

            "Farmer: " + name + "\n" +

            "Token: #" + token + "\n" +

            "Centre: " + centre + "\n" +

            "Slot: " + slot
        );


        // Open queue page

        window.location.href = "queue.html";

});