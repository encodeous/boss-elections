/**
 * @OnlyCurrentDoc
 */

class Voter {
    studentId: string;
    // Map of <Position Name, List of Candidates>
    votes: Map<string, string[]>;
}

class Ranking{
    candidateName: string;
    votes: number;
}

/**
 * Called when the sheet is loaded.
 */
function onOpen() {
    var menu = [
        { name: 'Instructions', functionName: 'infoFunction' },
        { name: 'Count Ballots', functionName: 'countBallots' }
    ];
    SpreadsheetApp.getActive().addMenu('BOSS Election', menu);
}

/**
 * Displays help / instructions to users
 */
function infoFunction() {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Documentation is Available', 'A guide on how this script works and how it is set up is available at https://github.com/encodeous/boss-elections', ui.ButtonSet.OK);
}

/**
 * Counts and Processes all of the responses recorded in the form
 */
function countBallots(){
    var ui = SpreadsheetApp.getUi();
    let [votes, warnings] = getResponses();
    if(warnings.length != 0){
        ui.alert('Some errors have occurred while processing results.', warnings.join('\n') , ui.ButtonSet.OK)
    }
    generateReport(votes);
}

/**
 * Extracts the real name of each position that people vote for, instead of the prompt
 * @param field
 */
function parseField(field: string){
    if(field.includes('"')){
        const spl = field.split('"');
        return spl[1];
    }else{
        return field;
    }
}

/**
 * Checks if the row in the sheet is empty
 * @param row
 */
function checkRowEmpty(row: any[]) : boolean{
    if(row == undefined) return true;
    for(const k of row){
        if(k !== ''){
            return false;
        }
    }
    return true;
}

/**
 * Parses the YRDSB Student Id out of the email
 * @param email
 */
function emailParser(email: string) : string | null{
    try{
        return email.toLowerCase().match(/(\d+)@gapps.yrdsb.ca/)[1];
    }catch{
        return null;
    }
}

/**
 * Fetches a list of the Allowed Students from the sheet
 */
function getAllowedStudents() : Set<string>{
    const sheet = SpreadsheetApp.getActive().getSheetByName('Allowed Students');
    const data = sheet.getRange('A:A').getValues();
    const students = new Set<string>();
    let cidx = 1;
    while(true){
        if(data[cidx] != undefined && data[cidx][0] !== ''){
            students.add(String(data[cidx][0]));
            cidx++;
        }else{
            break;
        }
    }
    return students;
}

/**
 * Sort comparator to compare candidate rankings
 * @param a
 * @param b
 */
function compare(a : Ranking, b : Ranking) {
    if (a.votes < b.votes){
        return 1;
    }
    if (a.votes > b.votes){
        return -1;
    }
    return 0;
}

/**
 * Generates a report in the "Results" sheet with the voting data
 * @param data
 */
function generateReport(data: Voter[]){
    let sheet = SpreadsheetApp.getActive().getSheetByName('Results');
    if(sheet != null){
        sheet.clear();
    }else{
        sheet = SpreadsheetApp.getActive().insertSheet('Results');
    }

    // Aggregated result. Map of <position name, <candidate, votes>>
    let aggr = new Map<string, Map<string, number>>();
    for(let vote of data){
        for(let position of Array.from(vote.votes.keys())){
            let val = vote.votes.get(position);
            // Check if position exists. If not, create it
            if(!aggr.has(position)){
                aggr.set(position, new Map<string, number>())
            }
            for (let candidate of val) {
                // check if a position has a candidate
                if(!aggr.get(position).has(candidate)){
                    aggr.get(position).set(candidate, 0);
                }
                // increment the candidate's vote by 1
                aggr.get(position).set(candidate, aggr.get(position).get(candidate) + 1);
            }
        }
    }

    let idx = 1;
    for (let position of Array.from(aggr.keys())) {
        // gets the frequency map of the current position.
        let freq = aggr.get(position);
        let candidates: Ranking[] = [];
        for(let candidate of Array.from(freq.keys())){
            let votes = freq.get(candidate);
            let rnk = new Ranking();
            rnk.votes = votes;
            rnk.candidateName = candidate;
            candidates.push(rnk)
        }
        // sort candidates by # of votes
        candidates.sort(compare);
        let heading = sheet.getRange(1, idx, 1, 2);
        heading.merge();
        heading.setBackground('gray');
        heading.setFontColor('white')
        heading.setValue(position);
        let lLabel = sheet.getRange(2, idx);
        lLabel.setFontStyle('italic');
        lLabel.setValue('Candidate Name');
        let rLabel = sheet.getRange(2, idx+1);
        rLabel.setFontStyle('italic');
        rLabel.setValue('Votes');
        for(let i = 0; i < candidates.length; i++){
            sheet.getRange(i + 3, idx).setValue(candidates[i].candidateName);
            sheet.getRange(i + 3, idx+1).setValue(candidates[i].votes);
        }
        sheet.autoResizeColumn(idx);
        sheet.autoResizeColumn(idx+1);
        idx += 2;
    }
}

/**
 * Fetches and validates the responses
 */
function getResponses() : [votes: Voter[], warnings: string[]]{
    const sheet = SpreadsheetApp.getActive().getSheetByName('Responses');
    if(sheet == null){
        return null;
    }
    const keys = sheet.getRange('1:1').getValues()[0]);
    const keyMap = new Map<string, number>();
    for(let i = 0; i < keys.length; i++){
        keyMap.set(keys[i], i);
    }
    const users = new Map<string, Voter>()
    sheet.getRange('A:Z').clearFormat();
    const data = sheet.getRange('A:Z').getValues();
    const allowedStudents = getAllowedStudents();
    console.log(allowedStudents);
    const warnings = [];
    for(let i = 1; i < data.length; i++){
        if(checkRowEmpty(data[i])) continue;
        const row = data[i];
        const id = emailParser(row[keyMap.get('Email Address')]);
        if(id == null){
            sheet.getRange(`${i+1}:${i+1}`).setBackground('red');
            warnings.push(`Failed to parse GAPPS Id: ${row[1]} on row ${i+1}`);
            continue;
        }
        if(users.has(id)){
            sheet.getRange(`${i+1}:${i+1}`).setBackground('orange');
            warnings.push(`Duplicate GAPPS Id ${id} found on row ${i+1}`);
            continue;
        }
        if(!allowedStudents.has(id)){
            sheet.getRange(`${i+1}:${i+1}`).setBackground('red');
            continue;
        }
        const userObj = new Voter();
        userObj.studentId = id;
        // make a map of Map<position, voted-candidates[]>
        userObj.votes = new Map<string, string[]>();
        for(let col = 0; col < keys.length; col++){
            if(keys[col] == 'Timestamp' || keys[col] == 'Email Address' || keys[col] === '') continue;
            userObj.votes.set(parseField(keys[col]), row[col].split(', '));
        }
        users.set(id, userObj);
    }
    return [Array.from(users.values()), warnings];
}