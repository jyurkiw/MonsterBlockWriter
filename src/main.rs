extern crate argparse;

use argparse::{ArgumentParser, Store, StoreTrue};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};

// Excel position constants (row, column)
const NAME: (u32, u32) = (0, 4);
const SIZE: (u32, u32) = (9, 20);
const ALIGNMENT: (u32, u32) = (10, 20);
const TYPE: (u32, u32) = (11, 20);
const ARMOR_CLASS: (u32, u32) = (26, 17);
const ARMOR: (u32, u32) = (32, 14);
const HIT_POINTS: (u32, u32) = (1, 15);
const HIT_DICE: (u32, u32) = (1, 16);
const STRENGTH: (u32, u32) = (2, 20);
const STRENGTH_BONUS: (u32, u32) = (2, 21);
const DEXTERITY: (u32, u32) = (3, 20);
const DEXTERITY_BONUS: (u32, u32) = (3, 21);
const CONSTITUTION: (u32, u32) = (4, 20);
const CONSTITUTION_BONUS: (u32, u32) = (4, 21);
const INTELLIGENCE: (u32, u32) = (5, 20);
const INTELLIGENCE_BONUS: (u32, u32) = (5, 21);
const WISDOM: (u32, u32) = (6, 20);
const WISDOM_BONUS: (u32, u32) = (6, 21);
const CHARMISMA: (u32, u32) = (7, 20);
const CHARMISMA_BONUS: (u32, u32) = (7, 21);
const CONDITION_IMMUNITIES: (u32, u32) = (46, 7);
const DAMAGE_RESISTANCES: (u32, u32) = (43, 7);
const DAMAGE_IMMUNITIES: (u32, u32) = (44, 7);
const DAMAGE_VULNERABILITIES: (u32, u32) = (45, 7);
const SENSES: (u32, u32) = (49, 7);
const LANGUAGES: (u32, u32) = (48, 7);
const CHALLENGE_RATING: (u32, u32) = (21, 17);
const XP_VALUE: (u32, u32) = (1, 11);
const LEGENDARY_RESISTANCES: (u32, u32) = (13, 14);
const ABILITIES: (u32, u32) = (56, 5);
const NUM_ABILITIES: (u32, u32) = (55, 9);
const LEGENDARY_ABILITIES: (u32, u32) = (56, 11);
const NUM_LEGENDARY_ABILITIES: (u32, u32) = (55, 14);
const MYTHIC_ABILITIES: (u32, u32) = (56, 16);
const NUM_MYTHIC_ABILITIES: (u32, u32) = (55, 20);
const LAIR_ACTIONS: (u32, u32) = (71, 5);
const NUM_LAIR_ACTIONS: (u32, u32) = (70, 8);
const SAVES: (u32, u32) = (51, 7);
const SKILLS: (u32, u32) = (52, 7);
const SPEED: (u32, u32) = (53, 7);
const ATTACKS: (u32, u32) = (35, 18);

#[derive(Debug)]
struct Monster {
    mname: String,
    msize: String,
    mtype: String,
    malign: String,
    mac: String,
    marmor: String,
    mhp: String,
    mhd: String,
    mstr: String,
    mstrbns: String,
    mdex: String,
    mdexbns: String,
    mcon: String,
    mconbns: String,
    mint: String,
    mintbns: String,
    mwis: String,
    mwisbns: String,
    mcha: String,
    mchabns: String,
    mcdtnims: String,
    mdmgres: String,
    mdmgims: String,
    mdmgvuln: String,
    msenses: String,
    mlanguages: String,
    mcr: String,
    mxp: String,
    mlngresists: String,
    msaves: String,
    mskills: String,
    mspeed: String,
    mabilities: Vec<String>,
    mlegabilities: Vec<String>,
    mmythabilities: Vec<String>,
    mlairactions: Vec<String>,
    mattacks: Vec<MonsterAttack>,
}

impl Monster {
    fn new() -> Monster {
        Monster {
            mname: "".to_string(),
            msize: "".to_string(),
            mtype: "".to_string(),
            malign: "".to_string(),
            mac: "".to_string(),
            marmor: "".to_string(),
            mhp: "".to_string(),
            mhd: "".to_string(),
            mstr: "".to_string(),
            mstrbns: "".to_string(),
            mdex: "".to_string(),
            mdexbns: "".to_string(),
            mcon: "".to_string(),
            mconbns: "".to_string(),
            mint: "".to_string(),
            mintbns: "".to_string(),
            mwis: "".to_string(),
            mwisbns: "".to_string(),
            mcha: "".to_string(),
            mchabns: "".to_string(),
            mcdtnims: "".to_string(),
            mdmgres: "".to_string(),
            mdmgims: "".to_string(),
            mdmgvuln: "".to_string(),
            msenses: "".to_string(),
            mlanguages: "".to_string(),
            mcr: "".to_string(),
            mxp: "".to_string(),
            mlngresists: "0".to_string(),
            msaves: "".to_string(),
            mskills: "".to_string(),
            mspeed: "".to_string(),
            mabilities: vec!["".to_string()],
            mlegabilities: vec!["".to_string()],
            mmythabilities: vec!["".to_string()],
            mlairactions: vec!["".to_string()],
            mattacks: vec![],
        }
    }
}

#[derive(Debug)]
struct MonsterAttack {
    valid: bool,
    maname: String,
    matype: String,
    maab: String,
    mareach: String,
    madamagecode: String,
}

impl MonsterAttack {
    fn get(r: &Range<DataType>, pos: (u32, u32)) -> MonsterAttack {
        let mut ma = MonsterAttack {
            valid: false,
            maname: get_value(r, (pos.0, pos.1)),
            matype: get_value(r, (pos.0, pos.1 + 1)),
            maab: get_value(r, (pos.0, pos.1 + 2)),
            mareach: get_value(r, (pos.0, pos.1 + 3)),
            madamagecode: get_value(r, (pos.0, pos.1 + 5)),
        };
        if ma.madamagecode.len() > 0 {
            ma.valid = true;
        }
        ma
    }

    fn get_all(r: &Range<DataType>, pos: (u32, u32)) -> Vec<MonsterAttack> {
        let mut vma: Vec<MonsterAttack> = Vec::new();
        for rowidx in pos.0..=pos.0 + 7 {
            let ma = MonsterAttack::get(r, (rowidx, pos.1));
            if ma.valid {
                vma.push(ma);
            }
        }
        vma
    }
}

fn get_value(r: &Range<DataType>, pos: (u32, u32)) -> String {
    let v = r.get_value(pos).unwrap();
    println!("{:?}", v);

    if v.is_string() {
        v.to_string()
    } else if v.is_int() {
        v.get_int().unwrap().to_string()
    } else if v.is_float() {
        v.get_float().unwrap().to_string()
    } else if v.is_bool() {
        if v.get_bool().unwrap() {
            "true".to_string()
        } else {
            "false".to_string()
        }
    } else {
        "".to_string()
    }
}

fn get_vertical_values(r: &Range<DataType>, pos: (u32, u32), rows: u32) -> Vec<String> {
    let mut values: Vec<String> = Vec::new();

    for i in pos.0..pos.0 + rows {
        values.push(get_value(r, (i, pos.1)));
    }

    values
}

fn write_standard_monster(monster: Monster) {
    println!("___");
    println!("{{{{monster,frame");

    println!("## {}", monster.mname);
    println!("*{} {}, {}*", monster.msize, monster.mtype, monster.malign);
    println!("___");
    println!("**Armor Class** :: {} ({})", monster.mac, monster.marmor);
    println!("**Hit Points**  :: {} ({})", monster.mhp, monster.mhd);
    println!("**Speed**       :: {}", monster.mspeed);
    println!("___");
    println!("|  STR  |  DEX  |  CON  |  INT  |  WIS  |  CHA  |");
    println!("|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|");
    println!(
        "|{} ({})|{} ({})|{} ({})|{} ({})|{} ({})|{} ({})|",
        monster.mstr,
        monster.mstrbns,
        monster.mdex,
        monster.mdexbns,
        monster.mcon,
        monster.mconbns,
        monster.mint,
        monster.mintbns,
        monster.mwis,
        monster.mwisbns,
        monster.mcha,
        monster.mchabns
    );
    println!("___");
    if monster.msaves.len() > 0 {
        println!("**Saves** :: {}", monster.msaves);
    }
    if monster.mskills.len() > 0 {
        println!("**Skills** :: {}", monster.mskills);
    }
    if monster.mdmgres.len() > 0 {
        println!("**Damage Resistances** :: {}", monster.mdmgres);
    }
    if monster.mdmgims.len() > 0 {
        println!("**Damage Immunities** :: {}", monster.mdmgims);
    }
    if monster.mdmgvuln.len() > 0 {
        println!("**Damage Vulnerabilities** :: {}", monster.mdmgvuln);
    }
    if monster.msenses.len() > 0 {
        println!("**Senses** :: {}", monster.msenses);
    }
    if monster.mlanguages.len() > 0 {
        println!("**Languages** :: {}", monster.mlanguages);
    }
    println!("**Challenge** :: {} ({} XP)", monster.mcr, monster.mxp);
    println!("___");
    for ability in monster.mabilities {
        println!("**{}** :: stuff", ability);
        println!(":");
    }
    println!("### Actions");
    for ability in monster.mattacks {
        if ability.matype != "Spell" {
            println!(
                "**{}** :: *{} weapon attack:* +{} to hit, {}, one target. *Hit* {}",
                ability.maname, ability.matype, ability.maab, ability.mareach, ability.madamagecode
            );
        } else {
            println!(
                "**{}** :: spell_description (range: {})",
                ability.maname, ability.mareach
            );
        }
    }

    if monster.mmythabilities.len() > 0 {
        println!("### Mythic Actions");
        for ability in monster.mmythabilities {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }

    if monster.mlegabilities.len() > 0 {
        println!("### Legendary Actions");
        for ability in monster.mlegabilities {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }

    if monster.mlairactions.len() > 0 {
        println!("### Lair Actions");
        for ability in monster.mlairactions {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }
    println!("}}}}");
}

fn main() {
    let mut filename: String = "".to_string();
    let mut verbose = false;
    let mut monster: Monster = Monster::new();
    {
        // this block limits scope of borrows by ap.refer() method
        let mut ap = ArgumentParser::new();
        ap.set_description("Write a Homebrewery statblock from a monster excel file.");
        ap.refer(&mut filename)
            .add_option(&["-f", "--file"], Store, "Path to the excel file.");
        ap.refer(&mut verbose).add_option(
            &["-v", "--verbose"],
            StoreTrue,
            "Enable verbose output.",
        );
        ap.parse_args_or_exit();
    }

    let mut excel: Xlsx<_> = open_workbook(filename).unwrap();

    if let Some(Ok(r)) = excel.worksheet_range_at(0) {
        monster.mname = get_value(&r, NAME);
        monster.msize = get_value(&r, SIZE);
        monster.mtype = get_value(&r, TYPE);
        monster.malign = get_value(&r, ALIGNMENT);
        monster.mac = get_value(&r, ARMOR_CLASS);
        monster.marmor = get_value(&r, ARMOR);
        monster.mhp = get_value(&r, HIT_POINTS);
        monster.mhd = get_value(&r, HIT_DICE);
        monster.mstr = get_value(&r, STRENGTH);
        monster.mstrbns = get_value(&r, STRENGTH_BONUS);
        monster.mdex = get_value(&r, DEXTERITY);
        monster.mdexbns = get_value(&r, DEXTERITY_BONUS);
        monster.mcon = get_value(&r, CONSTITUTION);
        monster.mconbns = get_value(&r, CONSTITUTION_BONUS);
        monster.mint = get_value(&r, INTELLIGENCE);
        monster.mintbns = get_value(&r, INTELLIGENCE_BONUS);
        monster.mwis = get_value(&r, WISDOM);
        monster.mwisbns = get_value(&r, WISDOM_BONUS);
        monster.mcha = get_value(&r, CHARMISMA);
        monster.mchabns = get_value(&r, CHARMISMA_BONUS);
        monster.mcdtnims = get_value(&r, CONDITION_IMMUNITIES);
        monster.mdmgres = get_value(&r, DAMAGE_RESISTANCES);
        monster.mdmgims = get_value(&r, DAMAGE_IMMUNITIES);
        monster.mdmgvuln = get_value(&r, DAMAGE_VULNERABILITIES);
        monster.msenses = get_value(&r, SENSES);
        monster.mlanguages = get_value(&r, LANGUAGES);
        monster.mcr = get_value(&r, CHALLENGE_RATING);
        monster.mxp = get_value(&r, XP_VALUE);
        monster.mlngresists = get_value(&r, LEGENDARY_RESISTANCES);
        monster.msaves = get_value(&r, SAVES);
        monster.mskills = get_value(&r, SKILLS);
        monster.mspeed = get_value(&r, SPEED);
        monster.mattacks = MonsterAttack::get_all(&r, ATTACKS);

        // Get Action Names
        monster.mabilities = get_vertical_values(
            &r,
            ABILITIES,
            get_value(&r, NUM_ABILITIES).parse::<u32>().unwrap(),
        );
        monster.mlegabilities = get_vertical_values(
            &r,
            LEGENDARY_ABILITIES,
            get_value(&r, NUM_LEGENDARY_ABILITIES)
                .parse::<u32>()
                .unwrap(),
        );
        monster.mmythabilities = get_vertical_values(
            &r,
            MYTHIC_ABILITIES,
            get_value(&r, NUM_MYTHIC_ABILITIES).parse::<u32>().unwrap(),
        );
        monster.mlairactions = get_vertical_values(
            &r,
            LAIR_ACTIONS,
            get_value(&r, NUM_LAIR_ACTIONS).parse::<u32>().unwrap(),
        );
        if verbose {
            // Describe all gathered data.
            println!("{:?}", monster);
        }

        write_standard_monster(monster);
    }
}
