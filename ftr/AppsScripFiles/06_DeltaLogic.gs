 
/**
 * @file 06_DeltaLogic.gs
 * @description Contains the high-performance replacement functions for calculating deltas with enhanced logging.
 */

function rackDelta(ss) {
    Logger.log("--- Starting rackDelta calculation ---");
    console.time("rackDelta");
    const rackDataSheet = ss.getSheetByName("RackData");
    const previousDataSheet = ss.getSheetByName("PreviousRackData");
    const rackDeltaSheet = ss.getSheetByName("RackDelta");

    const rackDataValues = rackDataSheet.getDataRange().getValues();
    const previousDataValues = previousDataSheet.getDataRange().getValues();

    const previousDataMap = new Map(previousDataValues.slice(1).map(row => [row[0], row]));

    const headers = [
        "Rack Inventory", "Datacenter", "Room", "Bank", "Rack Identifier",
        "Usage", "Usage KW-1", "Max Usage", "Max Usage KW-1"
    ];

    const outputData = [headers];
    const backgroundColors = [];

    for (let i = 1; i < rackDataValues.length; i++) {
        const currentRow = rackDataValues[i];
        const rackId = currentRow[0];
        const previousRow = previousDataMap.get(rackId);

        const currentUsage = currentRow[7];
        const currentMax = currentRow[8];

        if (!previousRow) {
            outputData.push([
                rackId, currentRow[1], currentRow[2], currentRow[3], currentRow[4],
                currentUsage, "", currentMax, ""
            ]);
        } else {
            const prevUsage = previousRow[7];
            const prevMax = previousRow[8];
            if (currentUsage !== prevUsage || currentMax !== prevMax) {
                outputData.push([
                    rackId, currentRow[1], currentRow[2], currentRow[3], currentRow[4],
                    currentUsage, prevUsage, currentMax, prevMax
                ]);
            }
        }
    }

    for (let i = 1; i < outputData.length; i++) {
        const row = outputData[i];
        const rowColors = new Array(headers.length).fill(null);
        if (row[6] !== "" && row[5] > row[6]) rowColors[5] = RED_BACKGROUND;
        if (row[6] !== "" && row[5] < row[6]) rowColors[5] = GREEN_BACKGROUND;
        if (row[8] !== "" && row[7] > row[8]) rowColors[7] = RED_BACKGROUND;
        if (row[8] !== "" && row[7] < row[8]) rowColors[7] = GREEN_BACKGROUND;
        backgroundColors.push(rowColors);
    }

    rackDeltaSheet.clearContents();
    if (outputData.length > 1) {
        rackDeltaSheet.getRange(1, 1, outputData.length, headers.length).setValues(outputData);
        rackDeltaSheet.getRange(2, 1, backgroundColors.length, headers.length).setBackgrounds(backgroundColors);
        Logger.log(`RackDelta updated: ${outputData.length - 1} changes detected.`);
    } else {
        rackDeltaSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        rackDeltaSheet.getRange(2, 1).setValue("No changes detected.");
        Logger.log("RackDelta: No changes detected.");
    }
    console.timeEnd("rackDelta");
}

function pduDelta(ss) {
    Logger.log("--- Starting pduDelta calculation ---");
    console.time("pduDelta");
    const pduDataSheet = ss.getSheetByName("PDUData");
    const previousPDUDataSheet = ss.getSheetByName("PreviousPDUData");
    const pduDeltaSheet = ss.getSheetByName("PDUDelta");

    const pduDataValues = pduDataSheet.getDataRange().getValues();
    const previousPDUDataValues = previousPDUDataSheet.getDataRange().getValues();
    const previousPduMap = new Map(previousPDUDataValues.slice(1).map(row => [row[1], row]));

    const outputData = [];
    const breakerHeaders = ["Limit Main", "Usage Main", "Limit A", "Usage A", "Limit B", "Usage B", "Limit C", "Usage C", "Limit D", "Usage D", "Limit E", "Usage E", "Limit F", "Usage F", "Limit G", "Usage G", "Limit H", "Usage H", "Limit I", "Usage I", "Limit J", "Usage J", "Limit K", "Usage K", "Limit L", "Usage L"];

    for (let i = 1; i < pduDataValues.length; i++) {
        const currentRow = pduDataValues[i];
        const serial = currentRow[1];
        const previousRow = previousPduMap.get(serial);
        const changes = [];

        if (!previousRow) {
            for (let j = 3; j < 29; j++) {
                if (currentRow[j] !== "") {
                    changes.push({ breaker: breakerHeaders[j - 3], current: currentRow[j], prev: "-" });
                }
            }
        } else {
            for (let j = 3; j < 29; j++) {
                if (currentRow[j] !== previousRow[j]) {
                    changes.push({ breaker: breakerHeaders[j - 3], current: currentRow[j] || "-", prev: previousRow[j] || "-" });
                }
            }
        }

        if (changes.length > 0) {
            const baseInfo = [currentRow[0], currentRow[1], currentRow[2]];
            let dynamicRow = [];
            changes.forEach(change => { dynamicRow.push(change.breaker, change.current, change.prev); });
            outputData.push([...baseInfo, ...dynamicRow]);
        }
    }

    pduDeltaSheet.clear();
    const headers = ["Rack Inventory", "Serial", "Residence"];
    pduDeltaSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    if (outputData.length > 0) {
        let maxCols = 0;
        outputData.forEach(row => { if (row.length > maxCols) maxCols = row.length; });
        const dataRange = pduDeltaSheet.getRange(2, 1, outputData.length, maxCols);
        dataRange.setValues(outputData.map(row => row.concat(new Array(maxCols - row.length).fill(""))));
        Logger.log(`PDUDelta updated: ${outputData.length} PDUs with changes.`);
    } else {
        pduDeltaSheet.getRange(2, 1).setValue("No changes detected.");
        Logger.log("PDUDelta: No changes detected.");
    }
    console.timeEnd("pduDelta");
}

function cablingDelta(ss) {
    Logger.log("--- Starting cablingDelta calculation ---");
    console.time("cablingDelta");
    const cablingDataSheet = ss.getSheetByName("CablingData");
    const previousCablingDataSheet = ss.getSheetByName("PreviousCablingData");
    const cablingDeltaSheet = ss.getSheetByName("CablingDelta");

    const currentValues = cablingDataSheet.getDataRange().getValues();
    const previousValues = previousCablingDataSheet.getDataRange().getValues();

    const headers = currentValues[0];
    const currentMap = new Map(currentValues.slice(1).map(row => [row[0], row]));
    const previousMap = new Map(previousValues.slice(1).map(row => [row[0], row]));

    const outputData = [headers];
    const backgroundColors = [];

    currentMap.forEach((currentRow, key) => {
        const previousRow = previousMap.get(key);
        if (!previousRow) {
            outputData.push(currentRow);
            backgroundColors.push(new Array(headers.length).fill(GREEN_BACKGROUND));
        } else {
            if (JSON.stringify(currentRow) !== JSON.stringify(previousRow)) {
                outputData.push(currentRow);
                const rowColors = currentRow.map((cell, i) => (cell !== previousRow[i] ? RED_BACKGROUND : null));
                backgroundColors.push(rowColors);
            }
        }
    });

    previousMap.forEach((previousRow, key) => {
        if (!currentMap.has(key)) {
            outputData.push(previousRow);
            backgroundColors.push(new Array(headers.length).fill(RED_BACKGROUND));
        }
    });

    cablingDeltaSheet.clear();
    if (outputData.length > 1) {
        const dataOnly = outputData.slice(1);
        cablingDeltaSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        cablingDeltaSheet.getRange(2, 1, dataOnly.length, headers.length).setValues(dataOnly);
        cablingDeltaSheet.getRange(2, 1, backgroundColors.length, headers.length).setBackgrounds(backgroundColors);
        Logger.log(`CablingDelta updated: ${outputData.length - 1} changes detected.`);
    } else {
        cablingDeltaSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        cablingDeltaSheet.getRange(2, 1).setValue("No changes detected.");
        Logger.log("CablingDelta: No changes detected.");
    }
    console.timeEnd("cablingDelta");
}
