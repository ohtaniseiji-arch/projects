fn main() {
    let a: u8 = 200;
    let b: u8 = 10;

    // checked_*: 正常時は Some(結果)
    let checked = a.checked_add(b);
    println!("checked_add: {:?}", checked);

    // overflowing_*: (結果, オーバーフローしたか)
    let (value, overflowed) = a.overflowing_add(b);
    println!("overflowing_add: value={value}, overflowed={overflowed}");

    // wrapping_*: ラップアラウンド
    let wrapped = a.wrapping_add(b);
    println!("wrapping_add: {wrapped}");

    // saturating_*: 最小値/最大値で飽和
    let saturated = a.saturating_add(b);
    println!("saturating_add: {saturated}");

    let x: u8 = 0;
    println!("checked_sub: {:?}", x.checked_sub(1));
    println!("overflowing_sub: {:?}", x.overflowing_sub(1));
    println!("wrapping_sub: {}", x.wrapping_sub(1));
    println!("saturating_sub: {}", x.saturating_sub(1));
}
