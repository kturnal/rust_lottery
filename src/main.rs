use std::*;
use std::collections::VecDeque;
use calamine::{Reader, Xlsx, open_workbook};
use rand::thread_rng;
use rand::seq::SliceRandom;

fn main()
{
    println!("[Lunch Lottery] Press Any Key to Start:");

    let mut input_string = String::new();
    io::stdin().read_line(&mut input_string)
    	.ok()
        .expect("Failed to read line");

    const ROW_SIZE: usize = 60;

    let mut excel_xlsx: Xlsx<_> = open_workbook("rust_lottery.xlsx").unwrap();
    let mut excel_table = vec![vec![String::from(""); ROW_SIZE]; ROW_SIZE];
    let mut participants: Vec<String> = Vec::new();

    if let Some(Ok(r)) = excel_xlsx.worksheet_range("rust_lottery")
    {
        let mut i = 0;
        for row in r.rows() {
            //println!("ROW {:?}", row);
            excel_table[i].clear();
            for elem in row
            {
                excel_table[i].push(elem.to_string());
            }
             i += 1;
        }
    }

    if let Some(Ok(r)) = excel_xlsx.worksheet_range("ParticipantList")
    {
        let mut i = 1;
        for row in r.rows() 
        {
            let input = row[0].to_string();
            if !input.is_empty()
            {
                participants.push(input);
                println!("participant #{i} {:?}", row[0].to_string());
            }

            i += 1;
        }
        println!("# OF PARTICIPANTS #{}", participants.len());
    }

    participants.shuffle(&mut thread_rng());

    // Matching is done here.
    let matches = validate_matches(&participants, &excel_table, 0);
    //Print out matches.
    print_matches(&matches, true);

    //TODO fill excel table back up.

} 

fn validate_matches(participants : &Vec<String>, table : &Vec<Vec<String>>, depth : usize) -> Vec<String>{

    let mut i = 0;
    let length = participants.len();
    let mut incorrect_matches: Vec<String> = Vec::new();
    let mut return_vec: Vec<String> = Vec::new();

    println!("Validate matches |\tdepth {}", depth);
    while i < length {
        if length > i+1 {

            let first_pos = get_person_index(i, &participants, &table);
            let second_pos = get_person_index(i+1, &participants, &table);
            let already_matched = table[first_pos][second_pos].ne("");

            if already_matched
            {
                //println!("{} and {} matched already at {} !!!!!!!!!!", vect[i], vect[i+1], table[first_pos][second_pos]);
                incorrect_matches.push(participants[i].clone());
                incorrect_matches.push(participants[i+1].clone());

            }
            else{
                println!("[validate_matches] new match: {} and {} ", participants[i], participants[i+1]);
                return_vec.push(participants[i].clone());
                return_vec.push(participants[i+1].clone());
            }
        }
        else { // TODO check if algorithm works with odd number 
            let first_pos = get_person_index(i, &participants, &table);
            println!("[validate_matches] {} has no match! ", participants[first_pos]);
        }

        i += 2;
    }

    for it in incorrect_matches.iter(){
        println!("Failed matches: {}\t| depth {}", it, depth);
    }
    println!();

    if incorrect_matches.len() > 0 {

        if incorrect_matches.len() == 2 // TODO randomly get a pair and try to match them
        {
            // TODO FIX.
            
            println!("-----------------Two people remaining, get FIRST two people");
            // let first = return_vec.pop().unwrap();
            // let second = return_vec.pop().unwrap();
            let return_vec_copy = return_vec.clone();
            let mut deque = VecDeque::from(return_vec_copy);
            let deque_first = deque.pop_back().unwrap();
            let deque_second = deque.pop_back().unwrap();
            println!("----lucky people {} {}", deque_first, deque_second);
    
            incorrect_matches.push(deque_first);
            incorrect_matches.push(deque_second);
        }

        incorrect_matches.shuffle(&mut thread_rng());
        if depth < 6
        {
            let mut new_matches = validate_matches(&incorrect_matches, &table, depth+1);
            println!("New matches:");
            print_matches(&new_matches, false);

            return_vec.append(&mut new_matches);
        }
        else
        {
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
/// # Examples
///
/// ```
/// // You can have rust code between fences inside the comments
/// // If you pass --test to `rustdoc`, it will even test it for you!
fn get_person_index(index : usize, vect : &Vec<String>, table : &Vec<Vec<String>>)  -> usize {
    let person_from_list = &vect[index];
    let mut table_index = 0;
    
    for elem in table[0].iter()
    {
        if person_from_list.eq(elem){
            //println!("#{} {} found on table index {}", index, person_from_list, table_index+1);
            return table_index;
        }
        table_index += 1;
    }
    println!("[get_person_index] Index not found for {} at index {} ", person_from_list, index);
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
