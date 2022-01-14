use std::collections::BTreeMap;

#[derive(Debug)]
pub struct StringStorage {
    pub index: BTreeMap<String, usize>,
    // Total count of strings (including doubles).
    pub total: usize,
    //  Size of the SST (count of unique strings).
    pub size: usize,
}

impl StringStorage {
    pub fn new() -> Self {
        StringStorage {
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
