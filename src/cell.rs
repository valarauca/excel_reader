
use super::calamine::{DataType,Range,Rows};
use std::fmt;


///Wraps DataType
pub struct Cell<'a> {
    x: &'a DataType
}
impl<'a> fmt::Debug for Cell<'a> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self.x {
            &DataType::Empty => write!(f, ","),
            &DataType::Bool(true) => write!(f,"{:?}" , true),
            &DataType::Bool(false) => write!(f,"{:?}" , false),
            &DataType::Int(i) => write!(f, "{:?},", i),
            &DataType::Float(x) => write!(f, "{:?},", x),
            &DataType::String(ref i) => write!(f, "{:?},", i),
            &DataType::Error(_) => write!(f, "CELLERROR,")
        }
    }
}
impl<'a> From<&'a DataType> for Cell<'a> {
    fn from(x: &'a DataType) -> Self {
        Cell{ x: x }
    }
}

///Wraps Rows
pub struct Row<'a> {
    data: &'a [DataType]
}
impl<'a> fmt::Debug for Row<'a> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        for item in self.data.iter().map(Cell::from) {
            write!(f, "{:?}", item);
        }
        write!(f,"\n")
    }
}
impl<'a> From<&'a [DataType]> for Row<'a> {
    fn from(x: &'a [DataType]) -> Self {
        Row{ data: x }
    }
}
pub struct Sheet<'a> {
    data: &'a Range
}
impl<'a> From<&'a Range> for Sheet<'a> {
    fn from(x: &'a Range) -> Self {
        Sheet{data: x}
    }
}
impl<'a> fmt::Debug for Sheet<'a> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        for row in self.data.rows().map(Row::from) {
            write!(f, "{:?}", row);
        }
        Ok(())
    }
}

pub fn to_csv<'a>(x: &'a Range) -> String {
    format!("{:?}", Sheet::from(x))
}
