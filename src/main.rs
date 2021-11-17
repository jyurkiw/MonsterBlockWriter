extern crate argparse;

use argparse::{ArgumentParser, Store, StoreTrue};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};

mod constants;

#[derive(Debug)]
struct Monster {
    mname: String,
    msize: String,
    mtype: String,
    malign: String,
    mac: String,
    marmor: String,
    mdc: String,
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
            mdc: "".to_string(),
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
    masavedc: String,
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
            masavedc: get_value(r, constants::SAVE_DC),
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

fn write_monster(monster: Monster, wide: bool) {
    println!("___");
    if wide {
        println!("{{{{monster,frame,wide");
    } else {
        println!("{{{{monster,frame");
    }

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
    println!(":");
    for ability in monster.mattacks {
        if ability.matype != "Spell" {
            println!(
                "**{}** :: *{} weapon attack:* +{} to hit, {}, one target. *Hit* {}",
                ability.maname, ability.matype, ability.maab, ability.mareach, ability.madamagecode
            );
            println!(":");
        } else {
            println!(
                "**{}** :: spell_description (range: {}, damage: {}, type: {}, dc: {})",
                ability.maname,
                ability.mareach,
                ability.madamagecode,
                ability.matype,
                ability.masavedc
            );
            println!(":");
        }
    }

    if monster.mmythabilities.len() > 0 {
        println!("### Mythic Actions");
        println!(":");
        for ability in monster.mmythabilities {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }

    if monster.mlegabilities.len() > 0 {
        println!("### Legendary Actions");
        println!(":");
        for ability in monster.mlegabilities {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }

    if monster.mlairactions.len() > 0 {
        println!("### Lair Actions");
        println!(":");
        for ability in monster.mlairactions {
            println!("**{}** :: stuff", ability);
            println!(":");
        }
    }
    println!("}}}}");
}

fn write_standard_monster(monster: Monster) {
    write_monster(monster, false);
}

fn write_wide_monster(monster: Monster) {
    write_monster(monster, true);
}

fn main() {
    let mut filename: String = "".to_string();
    let mut verbose = false;
    let mut wide = false;
    let mut monster: Monster = Monster::new();
    {
        // this block limits scope of borrows by ap.refer() method
        let mut ap = ArgumentParser::new();
        ap.set_description("Write a Homebrewery statblock from a monster excel file.");
        ap.refer(&mut filename)
            .add_option(&["-f", "--file"], Store, "Path to the excel file.");
        ap.refer(&mut wide)
            .add_option(&["-w", "--wide"], StoreTrue, "Print wide monster block.");
        ap.refer(&mut verbose).add_option(
            &["-v", "--verbose"],
            StoreTrue,
            "Enable verbose output.",
        );
        ap.parse_args_or_exit();
    }

    let mut excel: Xlsx<_> = open_workbook(filename).unwrap();

    if let Some(Ok(r)) = excel.worksheet_range_at(0) {
        monster.mname = get_value(&r, constants::NAME);
        monster.msize = get_value(&r, constants::SIZE);
        monster.mtype = get_value(&r, constants::TYPE);
        monster.malign = get_value(&r, constants::ALIGNMENT);
        monster.mac = get_value(&r, constants::ARMOR_CLASS);
        monster.marmor = get_value(&r, constants::ARMOR);
        monster.mhp = get_value(&r, constants::HIT_POINTS);
        monster.mhd = get_value(&r, constants::HIT_DICE);
        monster.mstr = get_value(&r, constants::STRENGTH);
        monster.mstrbns = get_value(&r, constants::STRENGTH_BONUS);
        monster.mdex = get_value(&r, constants::DEXTERITY);
        monster.mdexbns = get_value(&r, constants::DEXTERITY_BONUS);
        monster.mcon = get_value(&r, constants::CONSTITUTION);
        monster.mconbns = get_value(&r, constants::CONSTITUTION_BONUS);
        monster.mint = get_value(&r, constants::INTELLIGENCE);
        monster.mintbns = get_value(&r, constants::INTELLIGENCE_BONUS);
        monster.mwis = get_value(&r, constants::WISDOM);
        monster.mwisbns = get_value(&r, constants::WISDOM_BONUS);
        monster.mcha = get_value(&r, constants::CHARMISMA);
        monster.mchabns = get_value(&r, constants::CHARMISMA_BONUS);
        monster.mcdtnims = get_value(&r, constants::CONDITION_IMMUNITIES);
        monster.mdmgres = get_value(&r, constants::DAMAGE_RESISTANCES);
        monster.mdmgims = get_value(&r, constants::DAMAGE_IMMUNITIES);
        monster.mdmgvuln = get_value(&r, constants::DAMAGE_VULNERABILITIES);
        monster.msenses = get_value(&r, constants::SENSES);
        monster.mlanguages = get_value(&r, constants::LANGUAGES);
        monster.mcr = get_value(&r, constants::CHALLENGE_RATING);
        monster.mxp = get_value(&r, constants::XP_VALUE);
        monster.mlngresists = get_value(&r, constants::LEGENDARY_RESISTANCES);
        monster.msaves = get_value(&r, constants::SAVES);
        monster.mskills = get_value(&r, constants::SKILLS);
        monster.mspeed = get_value(&r, constants::SPEED);
        monster.mattacks = MonsterAttack::get_all(&r, constants::ATTACKS);

        // Get Action Names
        monster.mabilities = get_vertical_values(
            &r,
            constants::ABILITIES,
            get_value(&r, constants::NUM_ABILITIES)
                .parse::<u32>()
                .unwrap(),
        );
        monster.mlegabilities = get_vertical_values(
            &r,
            constants::LEGENDARY_ABILITIES,
            get_value(&r, constants::NUM_LEGENDARY_ABILITIES)
                .parse::<u32>()
                .unwrap(),
        );
        monster.mmythabilities = get_vertical_values(
            &r,
            constants::MYTHIC_ABILITIES,
            get_value(&r, constants::NUM_MYTHIC_ABILITIES)
                .parse::<u32>()
                .unwrap(),
        );
        monster.mlairactions = get_vertical_values(
            &r,
            constants::LAIR_ACTIONS,
            get_value(&r, constants::NUM_LAIR_ACTIONS)
                .parse::<u32>()
                .unwrap(),
        );
        if verbose {
            // Describe all gathered data.
            println!("{:?}", monster);
        }

        if wide {
            write_wide_monster(monster);
        } else {
            write_standard_monster(monster);
        }
    }
}
