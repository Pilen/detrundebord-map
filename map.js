
// using https://github.com/SheetJS/js-xlsx
// and google maps


var google = google;
var map;
var infowindow;
var data;
var headers;
var i = 0;
var places;
var over_query_limits = 0;
var currentInfoWindow;
var failed = [];

var query_failed_delay = 300;
var query_delay = 300;

function init() {
    $("#load").on("change", load);
    createMap();
}

function load(e) {
    var files = e.target.files;
    var file = files[0];
    var reader = new FileReader();
    reader.onerror = function (event) {
        console.log("Error while loading a file");
    };
    reader.onload = function(e) {
        var content = e.target.result;
        var workbook = XLSX.read(content, {type: "binary"});
        data = parseWorkbook(workbook);
        $("#form").hide();
        $("#map").css("visibility", "visible");

        populateMap();
    };
    reader.readAsBinaryString(file);
}

function parseWorkbook(workbook) {
    var firstSheetName = workbook.SheetNames[0];
    var sheet = workbook.Sheets[firstSheetName];
    var ref = sheet["!ref"];
    var match = (/:[^\d]+(\d+)/).exec(ref);
    var length = parseInt(match[1]);

    var data = [];

    headers = [];
    var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    for (var i = 0; i < alphabet.length; i++) {
        var header = getData(alphabet[i], 1);
        if (!header) {
            break;
        }
        headers.push(header);
    }
    // 1 indexed, first row is header
    for (var i = 2; i <= length; i++) {
        // Sted Adress Postnr By Kontaktnavn Email Telefon Levering åbent
        var place = {};
        place.name = getData("A", i);
        place.address = getData("B", i);
        place.zip = getData("C", i);
        place.city = getData("D", i);
        place.row = [];
        for (var j = 0; j < headers.length; j++) {
            place.row.push(getData(alphabet[j], i));
        }
        data.push(place);
    }
    return data;


    function getData(letter, number) {
        var value = sheet[letter + number];
        if (value) {
            return value.w;
        } else {
            return "";
        }
    }

}

function createMap(data) {
    var mapDiv = $("#map")[0];
    map = new google.maps.Map(mapDiv,
                              {center: {lat: 55.6901746,
                                        lng: 12.4818516},
                               zoom: 12});
    places = new google.maps.places.PlacesService(map);
}

function populateMap() {
    // if (i >= 5) {done(); return;}
    if (i >= data.length) {
        done();
        return;
    }

    var place = data[i];
    if (!place.address) {
        i++;
        populateMap();
        return;
    }
    var query = place.address + ", " + place.zip + " " + place.city;
    places.textSearch({query: query}, addMarker);

    function addMarker(results, status) {
        if (status === google.maps.places.PlacesServiceStatus.OVER_QUERY_LIMIT) {
            console.log("Over query limit");
            over_query_limits++;
            window.setTimeout(populateMap, query_failed_delay);
            return;
        }

        i++;
        if (status !== google.maps.places.PlacesServiceStatus.OK) {
            console.log("Error, status: ", status);
            failed.push("En fejl opstod da følgende skulle findes:", query);
            return;
        }
        if (results.length < 1) {
            failed.push("Følgende adresse kunne ikke findes:", query);

        }
        if (results.length > 1) {
            console.log("Følgende adresse gav flere ("+results.length+")resultater ", query);
            return;
        }

        // Sted Adress Postnr By Kontaktnavn Email Telefon Levering åbent
        var contentString = '<div class="content">' +
                '<h1>' + place.name + '</h1>' +
                '<p>' +
                place.address + ', ' + place.zip + ' ' + place.city +
                '</p>' +
                '<p>' +
                '<dl class="dl-horizontal">' +
                headers.map(function(header, i) {
                    return '<dt>' + header + '</dt>'+
                        '<dd>' + place.row[i] + '</dd>';
                }).join("") +
                '</dl>'+
                '</p>' +
                '</div>';
        var infoWindow = new google.maps.InfoWindow({
            content: contentString
        });
        var location = results[0];
        var marker = new google.maps.Marker({
            map: map,
            position: location.geometry.location,
            animation: google.maps.Animation.DROP,
            title: place.name,
            label: place.name
        });
        marker.addListener("click", function() {
            if (currentInfoWindow) {
                currentInfoWindow.close();
            }
            currentInfoWindow = infoWindow;
            infoWindow.open(map, marker);
        });

        window.setTimeout(populateMap, query_delay);
    }
}

function done() {
    var message;
    if (failed.length === 0) {
        message = "Alle steder er indlæst";
    } else {
        message = "<p>Der opstod fejl under indlæsningen</p>" + failed.map(function(error) {
            return "<p>"+error+"</p>";
        }).join("\n");
    }
    $("#status").html(message)
        .slideDown("fast", function() {
            var node = $(this);
            window.setTimeout(function () {
                node.fadeOut(1000);
            } , 1500);
        });
}

$(document).ready(init);
