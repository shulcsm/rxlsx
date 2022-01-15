use std::collections::BTreeMap;
use std::sync::{Arc, RwLock};

#[derive(Debug)]
pub struct Strings {
    pub index: BTreeMap<String, usize>,
    pub total: usize,
    pub size: usize,
}

impl Strings {
    pub fn new() -> Self {
        Strings {
            index: BTreeMap::new(),
            total: 0,
            size: 0,
        }
    }

    pub fn insert(&mut self, string: String) -> usize {
        self.total += 1;

        if let Some(idx) = self.index.get(&string) {
            *idx
        } else {
            let idx = self.size;
            self.index.insert(string, idx);
            self.size += 1;

            idx
        }
    }
}

#[derive(Debug)]
pub struct Shared {
    pub strings: Strings,
}

impl Shared {
    pub fn new() -> Self {
        Shared {
            strings: Strings::new(),
        }
    }
}

pub type SharedRef = Arc<RwLock<Shared>>;
