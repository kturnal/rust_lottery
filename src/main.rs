use std::*;
use std::collections::VecDeque;
use rand::thread_rng;
use rand::seq::SliceRandom;

use spreadsheet_ods::sheet::Range;
use spreadsheet_ods::{WorkBook, Sheet, Value};
use chrono::NaiveDate;
use spreadsheet_ods::format;
use spreadsheet_ods::formula;
use spreadsheet_ods::{cm, mm};
use spreadsheet_ods::style::{CellStyle};
use icu_locid::locale;
use spreadsheet_ods::style::units::{TextRelief, Border, Length};

const TABLE_SIZE: u32 = 49; //  length of row -1
const MAX_DEPTH: usize = 8; //    maximum number of traversals for the search
const DATE: &str = "99.99.24";

fn main()
{
    println!("[Lunch Lottery] Press Any Key to Start:");

    let mut input_string = String::new();
    io::stdin().read_line(&mut input_string)
    	.ok()
        .expect("Failed to read line");

    let path = std::path::Path::new("rust_lottery.ods");
    let mut wb = if path.exists() {
        spreadsheet_ods::read_ods(path).unwrap()
        } 
        else {
        WorkBook::new(locale!("en-US"))
    };

    // populate participants list.
    let partip_sheet = wb.sheet(6); // hard coded
    let mut participants: Vec<String> = Vec::new();

    for r in 0..TABLE_SIZE {
        let v = get_value(r, 0, partip_sheet).to_string();
        if !v.eq("")
        {
            participants.push(v);
        }
    }

    let main_sheet = wb.sheet(4); // hard coded
    let mut idx=0;

    for r in 0..=TABLE_SIZE{
        match main_sheet.value(r, 0) {
            spreadsheet_ods::Value::Empty => {}
            spreadsheet_ods::Value::Boolean(v) => {}
            spreadsheet_ods::Value::Number(v) => {}
            spreadsheet_ods::Value::Percentage(v) => {}
            spreadsheet_ods::Value::Currency(v, cur) => {}
            spreadsheet_ods::Value::Text(v) => {
                println!("[TEST] person_name:{} found index:{}", v, idx);
            }
            spreadsheet_ods::Value::TextXml(v) => {}
            spreadsheet_ods::Value::DateTime(v) => {}
            spreadsheet_ods::Value::TimeDuration(v) => {}
        }

        idx += 1;
    }

    let matches = matching(&participants, &main_sheet, 0);
    //validate matches
    assert!(matches.iter().all(|item| participants.contains(item)));
    //Print out matches.
    print_matches(&matches, true);

    for r in 0..TABLE_SIZE{
        for c in 0..TABLE_SIZE {
            // read value
            match main_sheet.value(r, c) {
                spreadsheet_ods::Value::Empty => {}
                spreadsheet_ods::Value::Boolean(v) => println!("({},{}) = bool {}", r, c, v),
                spreadsheet_ods::Value::Number(v) => println!("({},{}) = number {}", r, c, v),
                spreadsheet_ods::Value::Percentage(v) => println!("({},{}) = percent {}", r, c, v),
                spreadsheet_ods::Value::Currency(v, cur) => {
                    println!("({},{}) = currency {} {}", r, c, v, cur)
                }
                spreadsheet_ods::Value::Text(v) => {} //println!("({},{}) = text {}", r, c, v),
                spreadsheet_ods::Value::TextXml(v) => println!("({},{}) = xml {:?}", r, c, v),
                spreadsheet_ods::Value::DateTime(v) => println!("({},{}) = date {}", r, c, v),
                spreadsheet_ods::Value::TimeDuration(v) => {
                    println!("({},{}) = duration {}", r, c, v)
                }
            }
        }
    }

    //TODO update excel table.

    // let mut sheet = Sheet::new("sample");
    // sheet.set_value(2, 2, "sample");
    // wb.push_sheet(sheet);
    // spreadsheet_ods::write_ods( & mut wb, "tryout.ods");

} 

fn get_value(row : u32, col : u32, sheet : &Sheet ) -> &str {
        match sheet.value(row, col) {
            spreadsheet_ods::Value::Empty => {}
            spreadsheet_ods::Value::Boolean(v) => {}
            spreadsheet_ods::Value::Number(v) => {}
            spreadsheet_ods::Value::Percentage(v) => {}
            spreadsheet_ods::Value::Currency(v, cur) => {}
            spreadsheet_ods::Value::Text(v) => {
                return sheet.value(row, col).as_str_or("");
            }
            spreadsheet_ods::Value::TextXml(v) => {}
            spreadsheet_ods::Value::DateTime(v) => {}
            spreadsheet_ods::Value::TimeDuration(v) => {}
        }
        return "";
}

fn simple_match(participants : &Vec<String>, sheet : &Sheet) -> Vec<String> {

    let mut return_vec: Vec<String> = Vec::new();
    let mut incorrect_matches: Vec<String>;
    let mut partip_copy = participants.clone();
    let mut count : u32 = 0;

    // loop until a combination of matches is reached without any incorrect matches.
    loop
    {
        count +=1;
        let mut index = 0;
        incorrect_matches = Vec::new();
        return_vec = Vec::new();
        let length: usize = participants.len();
        println!("matching: |\tlength {} ", length);

        partip_copy.shuffle(&mut thread_rng());

        while index < length {
            if length > index+1 {
    
                let first = get_person_index(&partip_copy[index], sheet);
                let second = get_person_index(&partip_copy[index+1], sheet);
    
                let already_matched = get_value(first, second, sheet).ne("");
    
                //println!("{} and {} MATCH? {}.", &partip_copy[index], &partip_copy[index+1], get_value(first, second, sheet));
    
                if already_matched {
                    println!("{} and {} matched already at {}.", &partip_copy[index], &partip_copy[index+1], get_value(first, second, sheet));
                    incorrect_matches.push(partip_copy[index].clone());
                    incorrect_matches.push(partip_copy[index+1].clone());
    
                }
                else {
                    println!("[simple_match] #{} new match: {} and {} ", index, partip_copy[index], partip_copy[index+1]);
                    return_vec.push(partip_copy[index].clone());
                    return_vec.push(partip_copy[index+1].clone());
                }
            }
            else { // TODO make it work with odd number of people
                let index = get_person_index(&partip_copy[index], sheet);
                let index_usize : usize = index.try_into().unwrap();
                println!("!!!!!!!!!!!!!!!!!!!!!!!!!!![simple_match] {} has no match! ", partip_copy[index_usize]);
            }
    
            index += 2;
        }

        for it in incorrect_matches.iter() {
            println!("Failed to match: {}, count {}", it, count);
        }

        if incorrect_matches.is_empty()
        {
            println!("-------------------------------------------------------------no incorrect matches, return. Times tried:{}", count);
            break;
        }
        else {

            println!("--------------------------------------------------------------------Retry. Times tried:{}", count);
        }      
    }

    return return_vec;
}

fn matching(participants : &Vec<String>, sheet : &Sheet, depth: usize) -> Vec<String>{

    let length: usize = participants.len();
    let mut partip_copy = participants.clone();
    let mut index = 0;
    let mut incorrect_matches: Vec<String> = Vec::new();
    let mut return_vec: Vec<String> = Vec::new();

    println!("matching: |\tlength {} \tdepth {}", length, depth);
    partip_copy.shuffle(&mut thread_rng());

    while index < length {
        if length > index+1 {

            let first = get_person_index(&partip_copy[index], sheet);
            let second = get_person_index(&partip_copy[index+1], sheet);

            let already_matched = get_value(first, second, sheet).ne("");

            println!("{}:{} and {}:{} MATCH? {}.", &partip_copy[index], first, &partip_copy[index+1], second, get_value(first, second, sheet));

            if already_matched {
                //println!("{} and {} matched already at {}.", &participants[index], &participants[index+1], get_value(first, second, sheet));
                incorrect_matches.push(partip_copy[index].clone());
                incorrect_matches.push(partip_copy[index+1].clone());

            }
            else {
                println!("[validate_matches] new match: {} x {}", partip_copy[index], partip_copy[index+1]);
                return_vec.push(partip_copy[index].clone());
                return_vec.push(partip_copy[index+1].clone());
            }
        }
        else { // TODO check if algorithm works with odd number
            let index = get_person_index(&partip_copy[index], sheet);
            let index_usize : usize = index.try_into().unwrap();
            println!("[validate_matches] {} has no match! ", partip_copy[index_usize]);
        }

        index += 2;
    }

    if incorrect_matches.len() > 0 {

        for it in incorrect_matches.iter(){
            println!("Failed matches: {}\t| depth {}", it, depth);
        }

        if incorrect_matches.len() <= 4 // randomly add a pair to the pot and rematch
        {
            let mut vec_deque = VecDeque::from(return_vec.clone());
            let deque_first = vec_deque.pop_back().unwrap();
            let deque_second = vec_deque.pop_back().unwrap();
            println!("[Failed matches] less than 5 matches failed, move a match to that pot: {} x {}", deque_first, deque_second);
    
            incorrect_matches.push(deque_first);
            incorrect_matches.push(deque_second);

            return_vec.remove(return_vec.len()-1);
            return_vec.remove(return_vec.len()-1);
        }

        if depth < MAX_DEPTH {
            incorrect_matches.shuffle(&mut thread_rng());
            let mut new_matches = matching(&incorrect_matches, sheet, depth+1);
            return_vec.append(&mut new_matches);

        } else {
            println!("Depth exceeded {}, return.", MAX_DEPTH);
            for it in incorrect_matches.iter(){
                println!("Failed matches: {}\t| depth {}", it, depth);
            }
        }
    }
    else
    {
        println!("Everyone matched successfully.");
    }

    println!("Validation finished\t| depth {}", depth);

    return return_vec;
}

/// Returns the index of the employee on the Excel table, based on the index on the lottery participant list
///
/// # Arguments
///
/// * `person_name` - The name of the employee in participants list
/// * `sheet` - Sheet that has the employee match
fn get_person_index(person_name : &String, sheet : &Sheet)  -> u32 {

    let mut idx = 0;

    for r in 0..=TABLE_SIZE {
        let name: &str = get_value(r, 0, sheet);
        if person_name.eq(name)
        {
            //println!("[get_person_index] person_name:{} found index:{}", person_name, idx);
            return idx;
        }

        idx += 1;
    }

    println!("[get_person_index] Index not found for {}", person_name);
    return 0;
}

/// Prints the matches compatible to Slack
///
/// # Arguments
///
/// * `vect` - Final match list
fn print_matches(vect : &Vec<String>, extra_print : bool) {

    let mut i = 0;
    let size = vect.len();
    let mut match_count = 0;

    if extra_print
    {
        println!(" -- Final Result --");
    }

    while i < size {
        if size > (i+1) {
            println!("@{} and @{}", vect[i], vect[i+1]);
            match_count += 1;
        }
        else{
            println!("@{}", vect[i]);
        }
        i += 2;
    }

    if extra_print
    {
        println!("Succescully created {} matches from {} employees.", match_count, size);
    }
}