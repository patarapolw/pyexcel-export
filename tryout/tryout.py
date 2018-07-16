from pathlib import Path

if __name__ == '__main__':
    p = Path(__file__).parent.joinpath(Path('test.xs'))
    print(str(p.absolute()))
