use std::*;
use std::collections::VecDeque;
//use calamine::{Reader, Xlsx, open_workbook};
use rand::thread_rng;
use rand::seq::SliceRandom;


use spreadsheet_ods::sheet::Range;
use spreadsheet_ods::{WorkBook, Sheet, Value};
use chrono::NaiveDate;
use spreadsheet_ods::format;
use spreadsheet_ods::formula;
use spreadsheet_ods::{cm, mm};
use spreadsheet_ods::style::{CellStyle};
use spreadsheet_ods::color::Rgb;
use icu_locid::locale;
use spreadsheet_ods::style::units::{TextRelief, Border, Length};

const TABLE_SIZE: u32 = 9;

fn main()
{
    println!("[Lunch Lottery] Press Any Key to Start:");

    let mut input_string = String::new();
    io::stdin().read_line(&mut input_string)
    	.ok()
        .expect("Failed to read line");


    let loc = locale!("en-US"); 
    let path = std::path::Path::new("odsRustSheet.ods");
    let mut wb = if path.exists() {
        spreadsheet_ods::read_ods(path).unwrap()
        } 
        else {
        WorkBook::new(loc)
    };


    // populate participants list.
    let partip_sheet = wb.sheet(1);
    let mut participants: Vec<String> = Vec::new();
    for r in 0..TABLE_SIZE{
        match partip_sheet.value(r, 0) {
            spreadsheet_ods::Value::Empty => {}
            spreadsheet_ods::Value::Boolean(v) => {}
            spreadsheet_ods::Value::Number(v) => {}
            spreadsheet_ods::Value::Percentage(v) => {}
            spreadsheet_ods::Value::Currency(v, cur) => {}
            spreadsheet_ods::Value::Text(v) => {
                println!("({}) => {}", r, v);
                participants.push(v.to_string());
            }
            spreadsheet_ods::Value::TextXml(v) => {}
            spreadsheet_ods::Value::DateTime(v) => {}
            spreadsheet_ods::Value::TimeDuration(v) => {}
        }
    }

    let main_sheet = wb.sheet(0);
    //participants.shuffle(&mut thread_rng());
    let matches = matching(&participants, &main_sheet, 0);
    //Print out matches.
    print_matches(&matches, true);

    // let test_value = main_sheet.value(0, 2).as_str_or("");
    // println!("Test value: {}", &test_value);

    // let mut excel_xlsx: Xlsx<_> = open_workbook("rust_lottery.xlsx").unwrap();
    //let mut excel_table = vec![vec![String::from(""); ROW_SIZE]; ROW_SIZE];

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

    // get converted value
    let a1 = main_sheet.value(0, 0).as_str_or("");
    //println!("A1 {}", a1);

    // let mut sheet = Sheet::new("sample");
    // sheet.set_value(2, 2, "sample");
    // wb.push_sheet(sheet);
    // spreadsheet_ods::write_ods( & mut wb, "tryout.ods");

    // // Matching is done here.
    // let matches = validate_matches(&participants, &excel_table, 0);
    // //Print out matches.
    // print_matches(&matches, true);

    //TODO fill excel table back up.

} 

fn matching(participants : &Vec<String>, sheet : &Sheet, depth: usize) -> Vec<String>{

    let mut index = 0;
    let length: usize = participants.len();
    let mut incorrect_matches: Vec<String> = Vec::new();
    let mut return_vec: Vec<String> = Vec::new();

    println!("matching: |\tlength {} \tdepth {}", length, depth);

    // FIRST PART
    while index < length {
        if length > index+1 {

            let first_pos = get_person_index(&participants[index], sheet, TABLE_SIZE);
            let second_pos = get_person_index(&participants[index+1], sheet, TABLE_SIZE);

            let already_matched = sheet.value(first_pos, second_pos).as_str_or("").ne("");

            println!("{} and {} MATCH? {}.", &participants[index], &participants[index+1], sheet.value(first_pos, second_pos).as_str_or(""));

            if already_matched {
                println!("{} and {} matched already at {}.", &participants[index], &participants[index+1], sheet.value(first_pos, second_pos).as_str_or(""));
                incorrect_matches.push(participants[index].clone());
                incorrect_matches.push(participants[index+1].clone());

            }
            else {
                println!("{}[validate_matches] new match: {} and {} ", index, participants[index], participants[index+1]);
                return_vec.push(participants[index].clone());
                return_vec.push(participants[index+1].clone());
            }
        }
        else { // TODO check if algorithm works with odd number
            let first_pos = get_person_index(&participants[index], sheet, length.try_into().unwrap());
            let u32u : usize;
            u32u = first_pos.try_into().unwrap();
            println!("[validate_matches] {} has no match! ", participants[u32u]);
        }

        index += 2;
    }

    for it in incorrect_matches.iter(){
        println!("Failed matches: {}\t| depth {}", it, depth);
    }
    println!();

    if incorrect_matches.len() > 0 {

        if incorrect_matches.len() == 2 // TODO randomly get a pair and try to match them
        {
            // TODO FIX.
            
            println!("[Failed matches] Two people remaining, get LAST two people");
            // let first = return_vec.pop().unwrap();
            // let second = return_vec.pop().unwrap();
            let return_vec_copy = return_vec.clone();
            let mut deque = VecDeque::from(return_vec_copy);
            let deque_first = deque.pop_back().unwrap();
            let deque_second = deque.pop_back().unwrap();
            println!("----lucky people {} {}", deque_first, deque_second);
    
            incorrect_matches.push(deque_first);
            incorrect_matches.push(deque_second);

            return_vec.remove(return_vec.len()-1);
            return_vec.remove(return_vec.len()-2);
        }

        incorrect_matches.shuffle(&mut thread_rng());

        if depth < 6 {
            let mut new_matches = matching(&incorrect_matches, sheet, depth+1);
            println!("New matches:");
            return_vec.append(&mut new_matches);

        } else {
            println!("Depth exceeded 6, return.");
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
/// * `index` - Represents the index of the employee in `vect`
/// * `vect` - Vector of employees participating in lottery
/// * `table` - 2D Vector representing the Excel table
fn get_person_index(person_name : &String, sheet : &Sheet, number_participants : u32)  -> u32 {

    println!("[get_person_index_from_employee_list] person_name:{} number:{}", person_name, number_participants);
    let mut idx = 0;

    for r in 0..=number_participants{
        match sheet.value(r, 0) {
            spreadsheet_ods::Value::Empty => {}
            spreadsheet_ods::Value::Boolean(v) => {}
            spreadsheet_ods::Value::Number(v) => {}
            spreadsheet_ods::Value::Percentage(v) => {}
            spreadsheet_ods::Value::Currency(v, cur) => {}
            spreadsheet_ods::Value::Text(v) => {
                //println!("mainSheet index:{} value:{}", idx, v);
                if person_name.eq(v)
                {
                    println!("---{} found on table index {}", person_name, idx);
                    return idx;
                }
            }
            spreadsheet_ods::Value::TextXml(v) => {}
            spreadsheet_ods::Value::DateTime(v) => {}
            spreadsheet_ods::Value::TimeDuration(v) => {}
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